import json
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, colorchooser, filedialog
import sys
import time
import threading
import os
import logging
from datetime import datetime

# Global variables
refresh_thread = None
use_severity_colors = True
auto_scroll_enabled = True  # Auto-scroll defaults to ON
refresh_interval_ms = 3600000  # Default to 1 hour
next_sync_time = None  # To track next sync time
custom_severity_colors = {
    'low': '#00FF00',  # Green
    'medium': '#FFFF00',  # Yellow
    'high': '#FF0000'  # Red
}
# Store raw JSON data for each file with line mapping
raw_data_cache = {}

# Set up logging
def setup_logging():
    log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logviewer.log')
    logging.basicConfig(
        filename=log_file,
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    logging.info("Application started")

# Function to load JSON Lines data from file
def load_json(file_path, progress_callback=None, last_record_count=0, partial_load=False):
    logging.info(f"Loading JSON file: {file_path}, partial_load={partial_load}, last_record_count={last_record_count}")
    try:
        with open(file_path, 'r') as f:
            content = f.read().strip()
            lines = content.splitlines()
            total_lines = len(lines)
            
            if partial_load and last_record_count > 0:
                # Only process new lines
                new_lines = lines[last_record_count:]
                data = []
                for i, line in enumerate(new_lines):
                    line = line.strip()
                    if line:
                        item = json.loads(line)
                        if isinstance(item, dict):
                            data.append(item)
                    if progress_callback and len(new_lines) > 0:
                        progress_callback((i + 1) / len(new_lines))
                logging.info(f"Loaded {len(data)} new records from {file_path} (JSON Lines, partial)")
                if file_path in raw_data_cache:
                    raw_data_cache[file_path].extend(data)
                else:
                    raw_data_cache[file_path] = data
                return data, len(lines)
            
            # Full load
            try:
                data = json.loads(content)
                if not isinstance(data, list):
                    if isinstance(data, dict):
                        data = [data]  # Wrap single dict in a list
                    else:
                        raise ValueError("JSON must be a list of dictionaries or a single dictionary.")
                if not all(isinstance(item, dict) for item in data):
                    raise ValueError("JSON items must be dictionaries.")
                logging.info(f"Successfully loaded {len(data)} records from {file_path}")
                if progress_callback:
                    progress_callback(1.0)  # Signal completion
                raw_data_cache[file_path] = data  # Cache raw JSON data
                return data, len(lines)
            except json.JSONDecodeError:
                # Handle JSON Lines format
                data = []
                for i, line in enumerate(lines):
                    line = line.strip()
                    if line:
                        item = json.loads(line)
                        if isinstance(item, dict):
                            data.append(item)
                    if progress_callback and total_lines > 0:
                        progress_callback((i + 1) / total_lines)
                if not data:
                    raise ValueError("No valid JSON objects found.")
                logging.info(f"Successfully loaded {len(data)} records from {file_path} (JSON Lines)")
                raw_data_cache[file_path] = data  # Cache raw JSON data
                return data, len(lines)
    except Exception as e:
        logging.error(f"Failed to load JSON from {file_path}: {str(e)}")
        raise ValueError(f"Error loading JSON: {str(e)}")

# Function to flatten nested dictionaries for display
def flatten_dict(d, parent_key='', sep='.'):
    items = []
    for k, v in d.items():
        new_key = f"{parent_key}{sep}{k}" if parent_key else k
        if isinstance(v, dict):
            items.extend(flatten_dict(v, new_key, sep).items())
        else:
            items.append((new_key, v))
    return dict(items)

# Function to filter data based on column filters
def filter_data(data, filters, desired_columns):
    filtered_data = []
    for item in data:
        match = True
        for col, filter_text in filters.items():
            if filter_text:
                value = str(item.get(col, '')).lower()
                if filter_text.lower() not in value:
                    match = False
                    break
        if match:
            filtered_data.append(item)
    return filtered_data

# Function to format time remaining
def format_time_remaining(seconds):
    if seconds <= 0:
        return "Syncing now..."
    minutes, secs = divmod(int(seconds), 60)
    hours, minutes = divmod(minutes, 60)
    if hours > 0:
        return f"{hours}h {minutes}m {secs}s"
    elif minutes > 0:
        return f"{minutes}m {secs}s"
    else:
        return f"{secs}s"

# Function to update countdown timer for all tabs
def update_countdown_timer(tabs, desired_columns, root):
    global next_sync_time, refresh_interval_ms
    while True:
        if next_sync_time:
            seconds_left = max(0, (next_sync_time - time.time()))
            countdown_text = format_time_remaining(seconds_left)
            for file_path, (tree, error_label, filters, group_colors, last_manual_sync, last_auto_sync, record_count, json_text) in tabs.items():
                try:
                    # Get total records without loading full data
                    with open(file_path, 'r') as f:
                        total_lines = len(f.read().strip().splitlines())
                    status_text = f"Data loaded successfully. {len(tree.get_children())} of {total_lines} records displayed."
                    if last_manual_sync:
                        status_text += f"  Last Manual Sync: {last_manual_sync}"
                    if last_auto_sync:
                        status_text += f"  Last Automatic Sync: {last_auto_sync}"
                    status_text += f"  Next Automatic Sync: {countdown_text}"
                    error_label.config(text=status_text, justify='center', padx=20)
                    root.update_idletasks()
                except Exception as e:
                    logging.error(f"Failed to update countdown for {file_path}: {str(e)}")
        time.sleep(1)  # Update every second

# Function to refresh a single tab's table
def refresh_table(tree, file_path, error_label, filters, group_colors, desired_columns, json_text, is_auto_refresh=False, last_manual_sync=None, last_auto_sync=None, popup=None, record_count=0):
    global use_severity_colors, custom_severity_colors, auto_scroll_enabled, next_sync_time
    logging.info(f"Refreshing table for {file_path} (auto_refresh={is_auto_refresh}, record_count={record_count})")
    
    try:
        # Define progress callback for load_json
        def progress_callback(progress):
            if popup:
                popup.update_progress(progress)
                popup.top.update_idletasks()
        
        # Load data (partial for auto-refresh, full for manual/initial)
        data, new_record_count = load_json(file_path, progress_callback, record_count, partial_load=is_auto_refresh)
        if popup:
            popup.close()  # Close popup as soon as data is loaded
        if not data and not is_auto_refresh:
            error_label.config(text="No data to display.")
            logging.warning(f"No data found in {file_path}")
            return last_manual_sync, last_auto_sync, record_count
        
        # Flatten nested dictionaries
        flattened_data = [flatten_dict(item) for item in data]
        
        # Apply filters
        filtered_data = filter_data(flattened_data, filters, desired_columns)
        
        if not is_auto_refresh:
            # Clear existing items for full refresh
            for item in tree.get_children():
                tree.delete(item)
        
        # Set up columns (only for full refresh)
        if not is_auto_refresh:
            tree['columns'] = ['Line'] + desired_columns
            tree.heading('Line', text='Line', anchor='w')
            tree.column('Line', width=60, anchor='w', stretch=False)
            for col in desired_columns:
                tree.heading(col, text=col, anchor='w')
                tree.column(col, width=150, anchor='w', stretch=True)
        
        # Insert rows with line numbers and apply colors
        start_idx = len(tree.get_children()) + 1 if is_auto_refresh else 1
        for idx, item in enumerate(filtered_data, start_idx):
            group = str(item.get('group', ''))
            severity = str(item.get('severity', '')).lower()
            # Prioritize group color, fall back to severity color if enabled
            if group in group_colors:
                tag = f"group_{group}"
            elif use_severity_colors and severity in custom_severity_colors:
                tag = f"severity_{severity}"
            else:
                tag = ""
            values = [str(idx)] + [str(item.get(col, '')) for col in desired_columns]
            tree.insert('', 'end', values=values, tags=(tag,))
        
        # Apply colors to tags
        for group, color in group_colors.items():
            tree.tag_configure(f"group_{group}", background=color)
            logging.info(f"Applied group color for {group}: {color}")
        if use_severity_colors:
            for severity, color in custom_severity_colors.items():
                tree.tag_configure(f"severity_{severity}", background=color)
                logging.info(f"Applied severity color for {severity}: {color}")
        
        # Scroll to the bottom if auto-scroll is enabled
        if auto_scroll_enabled:
            tree.yview_moveto(1.0)  # Scroll to the bottom
        
        # Update sync timestamps
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if is_auto_refresh:
            last_auto_sync = current_time
            next_sync_time = time.time() + (refresh_interval_ms / 1000)
        else:
            last_manual_sync = current_time
        
        # Display with even spacing, including countdown
        total_records = new_record_count
        status_text = f"Data loaded successfully. {len(tree.get_children())} of {total_records} records displayed."
        if last_manual_sync:
            status_text += f"  Last Manual Sync: {last_manual_sync}"
        if last_auto_sync:
            status_text += f"  Last Automatic Sync: {last_auto_sync}"
        if next_sync_time:
            seconds_left = max(0, (next_sync_time - time.time()))
            status_text += f"  Next Automatic Sync: {format_time_remaining(seconds_left)}"
        error_label.config(text=status_text, justify='center', padx=20)
        logging.info(f"Table refreshed for {file_path}: {len(tree.get_children())} of {total_records} records displayed, manual={last_manual_sync}, auto={last_auto_sync}")
        
        return last_manual_sync, last_auto_sync, new_record_count
    except Exception as e:
        error_label.config(text=f"Error: {str(e)}")
        messagebox.showerror("Error", f"Failed to load JSON from {file_path}: {str(e)}")
        logging.error(f"Failed to refresh table for {file_path}: {str(e)}")
        return last_manual_sync, last_auto_sync, record_count

# Class for Please Wait popup with progress bar
class PleaseWaitPopup:
    def __init__(self, parent):
        self.top = tk.Toplevel(parent)
        self.top.transient(parent)
        self.top.grab_set()
        self.top.title("Loading")
        tk.Label(self.top, text="Loading data...", padx=20, pady=10).pack()
        self.progress = ttk.Progressbar(self.top, orient="horizontal", length=200, mode="determinate")
        self.progress.pack(pady=10)
        self.top.geometry("250x100")
        # Center the popup
        self.top.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.top.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.top.winfo_height()) // 2
        self.top.geometry(f"+{x}+{y}")
        self.top.resizable(False, False)
        logging.info("Please Wait popup opened with progress bar")

    def update_progress(self, value):
        self.progress['value'] = value * 100  # Convert to percentage
        self.top.update_idletasks()

    def close(self):
        self.top.grab_release()
        self.top.destroy()
        logging.info("Please Wait popup closed")

# Function to create context menu for column headers
def create_context_menu(tree, file_path, error_label, filters, group_colors, desired_columns):
    menu = tk.Menu(tree, tearoff=0)
    
    def show_filter_dialog(column):
        current_filter = filters.get(column, '')
        new_filter = simpledialog.askstring(
            "Filter", f"Enter filter for {column} (leave empty to clear):", initialvalue=current_filter, parent=tree
        )
        if new_filter is not None:  # None if dialog is canceled
            filters[column] = new_filter
            logging.info(f"Filter set for {column}: {new_filter}")
            popup = PleaseWaitPopup(tree.winfo_toplevel())
            tabs[file_path] = refresh_table(
                tree, file_path, error_label, filters, group_colors, desired_columns, tabs[file_path][7],
                popup=popup, record_count=tabs[file_path][6]
            ) + (tabs[file_path][6], tabs[file_path][7])
    
    def show_color_dialog():
        group_value = simpledialog.askstring(
            "Group Color", "Enter group value to color (e.g., AD_SYNC):", parent=tree
        )
        if group_value:
            color = colorchooser.askcolor(title=f"Choose color for group {group_value}", parent=tree)
            if color[1]:  # color[1] is the hex code, None if canceled
                group_colors[group_value] = color[1]
                logging.info(f"Color set for group {group_value} in {file_path}: {color[1]}")
                popup = PleaseWaitPopup(tree.winfo_toplevel())
                tabs[file_path] = refresh_table(
                    tree, file_path, error_label, filters, group_colors, desired_columns, tabs[file_path][7],
                    popup=popup, record_count=tabs[file_path][6]
                ) + (tabs[file_path][6], tabs[file_path][7])
    
    for col in desired_columns:
        menu.add_command(label=f"Filter {col}", command=lambda c=col: show_filter_dialog(c))
        if col == 'group':
            menu.add_command(label="Set Color for Group", command=show_color_dialog)
    
    def on_right_click(event):
        # Identify the column clicked
        col = tree.identify_column(event.x)
        if col != '#1':  # Skip Line column
            col_name = tree['columns'][int(col[1:])]  # Convert #2 to index 1, etc.
            if col_name in desired_columns:
                menu.post(event.x_root, event.y_root)
    
    tree.bind('<Button-3>', on_right_click)
    return menu

# Function to handle row selection and display raw JSON
def on_row_select(event, tree, file_path, json_text):
    selection = tree.selection()
    if not selection:
        json_text.config(state='normal')
        json_text.delete(1.0, tk.END)
        json_text.config(state='disabled')
        return
    item = tree.item(selection[0])
    line_number = int(item['values'][0]) - 1  # Line number is 1-based in Treeview, 0-based in data
    if file_path in raw_data_cache and line_number < len(raw_data_cache[file_path]):
        raw_json = raw_data_cache[file_path][line_number]
        formatted_json = json.dumps(raw_json, indent=2)
        json_text.config(state='normal')
        json_text.delete(1.0, tk.END)
        json_text.insert(tk.END, formatted_json)
        json_text.config(state='disabled')
        logging.info(f"Displayed raw JSON for line {line_number + 1} in {file_path}")
    else:
        json_text.config(state='normal')
        json_text.delete(1.0, tk.END)
        json_text.insert(tk.END, "Raw JSON data not available.")
        json_text.config(state='disabled')
        logging.warning(f"Raw JSON not found for line {line_number + 1} in {file_path}")

# Function to toggle severity colors
def toggle_severity_colors(tabs, desired_columns, root):
    global use_severity_colors
    use_severity_colors = not use_severity_colors
    status = "enabled" if use_severity_colors else "disabled"
    logging.info(f"Severity colors {status}")
    manual_refresh(tabs, desired_columns, root)

# Function to toggle auto-scroll
def toggle_auto_scroll(tabs, desired_columns, root):
    global auto_scroll_enabled
    auto_scroll_enabled = not auto_scroll_enabled
    status = "enabled" if auto_scroll_enabled else "disabled"
    logging.info(f"Auto-scroll {status}")
    if auto_scroll_enabled:
        # Scroll all tabs to the bottom
        for file_path, (tree, _, _, _, _, _, _, _) in tabs.items():
            tree.yview_moveto(1.0)
        root.update()

# Function to set custom severity colors
def set_custom_severity_colors(tabs, desired_columns, root):
    global custom_severity_colors
    for severity in custom_severity_colors:
        color = colorchooser.askcolor(title=f"Choose color for severity {severity.capitalize()}", parent=root)
        if color[1]:  # color[1] is the hex code, None if canceled
            custom_severity_colors[severity] = color[1]
            logging.info(f"Custom color set for severity {severity}: {color[1]}")
    manual_refresh(tabs, desired_columns, root)

# Function to set refresh interval
def set_refresh_interval(root, tabs, desired_columns):
    global refresh_thread, refresh_interval_ms, next_sync_time
    interval_minutes = simpledialog.askinteger(
        "Set Automatic Sync Interval", "Enter automatic sync interval in minutes (e.g., 30):", 
        parent=root, minvalue=1, initialvalue=60
    )
    if interval_minutes is None:  # Canceled
        logging.info("Set automatic sync interval dialog canceled")
        return
    
    logging.info(f"Set automatic sync interval to {interval_minutes} minutes")
    refresh_interval_ms = interval_minutes * 60 * 1000
    next_sync_time = time.time() + (refresh_interval_ms / 1000)
    
    # Start new refresh thread with updated interval
    periodic_refresh(tabs, desired_columns, root, refresh_interval_ms)

# Function to periodically refresh all tabs
def periodic_refresh(tabs, desired_columns, root, interval_ms):
    global refresh_thread, next_sync_time
    def run():
        # Start countdown timer thread
        countdown_thread = threading.Thread(target=update_countdown_timer, args=(tabs, desired_columns, root), daemon=True)
        countdown_thread.start()
        
        while True:
            for file_path, (tree, error_label, filters, group_colors, last_manual_sync, last_auto_sync, record_count, json_text) in list(tabs.items()):
                try:
                    popup = PleaseWaitPopup(root)
                    last_manual_sync, last_auto_sync, new_record_count = refresh_table(
                        tree, file_path, error_label, filters, group_colors, desired_columns, json_text,
                        is_auto_refresh=True, last_manual_sync=last_manual_sync, last_auto_sync=last_auto_sync, 
                        popup=popup, record_count=record_count
                    )
                    tabs[file_path] = (tree, error_label, filters, group_colors, last_manual_sync, last_auto_sync, new_record_count, json_text)
                    root.update()
                except Exception as e:
                    error_label.config(text=f"Error: {str(e)}")
                    messagebox.showerror("Error", f"Failed to load JSON from {file_path}: {str(e)}")
                    logging.error(f"Periodic refresh failed for {file_path}: {str(e)}")
            time.sleep(interval_ms / 1000)
    
    refresh_thread = threading.Thread(target=run, daemon=True)
    refresh_thread.start()
    logging.info(f"Periodic refresh started with interval {interval_ms/1000/60} minutes")

# Function to manually refresh all tabs
def manual_refresh(tabs, desired_columns, root):
    for file_path, (tree, error_label, filters, group_colors, last_manual_sync, last_auto_sync, record_count, json_text) in list(tabs.items()):
        try:
            popup = PleaseWaitPopup(root)
            last_manual_sync, last_auto_sync, new_record_count = refresh_table(
                tree, file_path, error_label, filters, group_colors, desired_columns, json_text,
                is_auto_refresh=False, last_manual_sync=last_manual_sync, last_auto_sync=last_auto_sync, 
                popup=popup, record_count=0  # Full refresh, reset record count
            )
            tabs[file_path] = (tree, error_label, filters, group_colors, last_manual_sync, last_auto_sync, new_record_count, json_text)
            root.update()
        except Exception as e:
            error_label.config(text=f"Error: {str(e)}")
            messagebox.showerror("Error", f"Failed to load JSON from {file_path}: {str(e)}")
            logging.error(f"Manual refresh failed for {file_path}: {str(e)}")

def main():
    setup_logging()
    logging.info("Starting main function")
    
    # Define the fields we care about
    desired_columns = ['source_info.ip', 'severity', 'type', 'name', 'id', 'group', 'rt', 'dhost', 'endpoint_id', 'endpoint_type']
    
    root = tk.Tk()
    root.title("Sophos SIEM Log Viewer")
    root.geometry("1200x600")  # Set initial window size
    logging.info("Main window created")
    
    # Create menu bar
    menubar = tk.Menu(root)
    file_menu = tk.Menu(menubar, tearoff=0)
    file_menu.add_command(label="Add File", command=lambda: add_file(root, notebook, tabs, desired_columns))
    menubar.add_cascade(label="File", menu=file_menu)
    
    options_menu = tk.Menu(menubar, tearoff=0)
    options_menu.add_command(label="Manual Sync", command=lambda: manual_refresh(tabs, desired_columns, root))
    options_menu.add_command(label="Set Automatic Sync Interval", command=lambda: set_refresh_interval(root, tabs, desired_columns))
    options_menu.add_command(label="Toggle Severity Colors", command=lambda: toggle_severity_colors(tabs, desired_columns, root))
    options_menu.add_command(label="Toggle Auto-Scroll", command=lambda: toggle_auto_scroll(tabs, desired_columns, root))
    options_menu.add_command(label="Set Custom Severity Colors", command=lambda: set_custom_severity_colors(tabs, desired_columns, root))
    menubar.add_cascade(label="Options", menu=options_menu)
    
    root.config(menu=menubar)
    logging.info("Menu bar configured with File and Options menus")
    
    # Create notebook for tabs
    notebook = ttk.Notebook(root)
    notebook.pack(expand=True, fill='both')
    logging.info("Notebook created")
    
    tabs = {}
    # Load files from command-line arguments or prompt for file
    file_paths = sys.argv[1:] if len(sys.argv) > 1 else []
    if not file_paths:
        file_path = filedialog.askopenfilename(
            title="Select Log File",
            filetypes=[("Log Files", "*.txt *.json *.jsonl *.cef"), ("All Files", "*.*")]
        )
        if file_path:
            file_paths = [file_path]
        else:
            logging.info("No file selected, starting with empty GUI")
    
    global next_sync_time
    next_sync_time = time.time() + (refresh_interval_ms / 1000)
    
    if file_paths:
        for file_path in file_paths:
            # Create tab
            tab = ttk.Frame(notebook)
            tab_name = os.path.basename(file_path).replace('.jsonl', '').replace('.json', '').replace('.txt', '').replace('.cef', '')
            notebook.add(tab, text=tab_name)
            logging.info(f"Tab created for {file_path}")
            
            # Error label
            error_label = tk.Label(tab, text="", fg="red", wraplength=1100)
            error_label.pack(pady=5, fill='x')
            
            # Create Treeview with scrollbars
            frame = tk.Frame(tab)
            frame.pack(expand=True, fill='both')
            
            tree = ttk.Treeview(frame, show='headings')
            vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
            hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
            tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
            
            vsb.pack(side='right', fill='y')
            hsb.pack(side='bottom', fill='x')
            tree.pack(expand=True, fill='both')
            logging.info(f"Treeview and scrollbars created for {file_path}")
            
            # Create Text widget for raw JSON display
            json_text = tk.Text(tab, height=10, wrap='word', state='disabled')
            json_text.pack(pady=5, fill='x')
            json_text_scroll = ttk.Scrollbar(tab, orient="vertical", command=json_text.yview)
            json_text_scroll.pack(side='right', fill='y')
            json_text.configure(yscrollcommand=json_text_scroll.set)
            logging.info(f"Raw JSON Text widget created for {file_path}")
            
            # Bind row selection to display raw JSON
            tree.bind('<<TreeviewSelect>>', lambda e: on_row_select(e, tree, file_path, json_text))
            
            # Initialize filters, group colors, and record count for this tab
            filters = {col: '' for col in desired_columns}
            group_colors = {}
            last_manual_sync = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            last_auto_sync = None
            record_count = 0
            
            # Create context menu for column headers
            create_context_menu(tree, file_path, error_label, filters, group_colors, desired_columns)
            
            # Initial load
            try:
                popup = PleaseWaitPopup(root)
                last_manual_sync, last_auto_sync, record_count = refresh_table(
                    tree, file_path, error_label, filters, group_colors, desired_columns, json_text,
                    is_auto_refresh=False, last_manual_sync=last_manual_sync, last_auto_sync=last_auto_sync, 
                    popup=popup, record_count=record_count
                )
                tabs[file_path] = (tree, error_label, filters, group_colors, last_manual_sync, last_auto_sync, record_count, json_text)
                root.update()
            except Exception as e:
                error_label.config(text=f"Error: {str(e)}")
                messagebox.showerror("Error", f"Failed to load JSON from {file_path}: {str(e)}")
                logging.error(f"Failed to load file {file_path}: {str(e)}")
            
            # Store tab components
            tabs[file_path] = (tree, error_label, filters, group_colors, last_manual_sync, last_auto_sync, record_count, json_text)
            logging.info(f"Tab components stored for {file_path}")
    
    # Start periodic refresh with default interval
    periodic_refresh(tabs, desired_columns, root, refresh_interval_ms)
    
    try:
        root.mainloop()
    except Exception as e:
        logging.error(f"Main loop crashed: {str(e)}")
        raise

def add_file(root, notebook, tabs, desired_columns):
    global next_sync_time
    file_path = filedialog.askopenfilename(
        title="Select Log File",
        filetypes=[("Log Files", "*.txt *.json *.jsonl *.cef"), ("All Files", "*.*")]
    )
    if file_path:
        tab = ttk.Frame(notebook)
        tab_name = os.path.basename(file_path).replace('.jsonl', '').replace('.json', '').replace('.txt', '').replace('.cef', '')
        notebook.add(tab, text=tab_name)
        logging.info(f"Tab created for {file_path}")
        
        # Error label
        error_label = tk.Label(tab, text="", fg="red", wraplength=1100)
        error_label.pack(pady=5, fill='x')
        
        # Create Treeview with scrollbars
        frame = tk.Frame(tab)
        frame.pack(expand=True, fill='both')
        
        tree = ttk.Treeview(frame, show='headings')
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        tree.pack(expand=True, fill='both')
        logging.info(f"Treeview and scrollbars created for {file_path}")
        
        # Create Text widget for raw JSON display
        json_text = tk.Text(tab, height=10, wrap='word', state='disabled')
        json_text.pack(pady=5, fill='x')
        json_text_scroll = ttk.Scrollbar(tab, orient="vertical", command=json_text.yview)
        json_text_scroll.pack(side='right', fill='y')
        json_text.configure(yscrollcommand=json_text_scroll.set)
        logging.info(f"Raw JSON Text widget created for {file_path}")
        
        # Bind row selection to display raw JSON
        tree.bind('<<TreeviewSelect>>', lambda e: on_row_select(e, tree, file_path, json_text))
        
        # Initialize filters, group colors, and record count for this tab
        filters = {col: '' for col in desired_columns}
        group_colors = {}
        last_manual_sync = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        last_auto_sync = None
        record_count = 0
        
        # Create context menu for column headers
        create_context_menu(tree, file_path, error_label, filters, group_colors, desired_columns)
        
        # Initial load
        try:
            popup = PleaseWaitPopup(root)
            last_manual_sync, last_auto_sync, record_count = refresh_table(
                tree, file_path, error_label, filters, group_colors, desired_columns, json_text,
                is_auto_refresh=False, last_manual_sync=last_manual_sync, last_auto_sync=last_auto_sync, 
                popup=popup, record_count=record_count
            )
            tabs[file_path] = (tree, error_label, filters, group_colors, last_manual_sync, last_auto_sync, record_count, json_text)
            root.update()
        except Exception as e:
            error_label.config(text=f"Error: {str(e)}")
            messagebox.showerror("Error", f"Failed to load JSON from {file_path}: {str(e)}")
            logging.error(f"Failed to load file {file_path}: {str(e)}")
        
        # Store tab components
        tabs[file_path] = (tree, error_label, filters, group_colors, last_manual_sync, last_auto_sync, record_count, json_text)
        logging.info(f"Tab components stored for {file_path}")

if __name__ == "__main__":
    main()