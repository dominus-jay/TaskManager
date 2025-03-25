#!/usr/bin/env pythonw
"""
Task Manager GUI - A graphical application to manage daily tasks.
"""
import os
import datetime
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from tkcalendar import DateEntry
from typing import Dict, List, Optional, Any, Tuple
import subprocess  # Add this import for the subprocess module
import webbrowser
import logging
import threading
import pymssql  # Changed: Import pymssql instead of pyodbc
import getpass # Import the 'getpass' module to get username
import tkinter.messagebox as messagebox # Import messagebox
import platform
import sys
import mysql.connector
import sqlite3

# Configure logging at the beginning of your script, before any other code
logging.basicConfig(
    filename='task_manager.log',  # Log file name
    level=logging.INFO,         # Log level: ERROR, WARNING, INFO, DEBUG
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# List of users that can be assigned tasks
USERS = ["All", "Jay", "Jude", "Jorgen", "Earl", "Philip", "Sam", "Glenn"]

class TaskManagerApp:
    """GUI application for managing tasks."""
    
    def __init__(self, root, task_manager):
        """Initialize the GUI application."""
        self.root = root
        self.root.title("Task Manager")
        self.root.geometry("900x600")
        self.root.minsize(800, 500)
        
        # Configure styles
        self.style = ttk.Style()
        self.style.configure("TButton", padding=6, font=('Helvetica', 10))
        self.style.configure("TLabel", font=('Helvetica', 10))
        self.style.configure("Header.TLabel", font=('Helvetica', 12, 'bold'))
        self.style.configure("Title.TLabel", font=('Helvetica', 16, 'bold'))
        
        # Try to set a modern theme if available
        try:
            self.root.tk.call("source", "azure.tcl")
            self.style.theme_use("azure")
        except tk.TclError:
            # If azure theme is not available, try other themes
            available_themes = self.style.theme_names()
            for theme in ["clam", "alt", "vista"]:
                if theme in available_themes:
                    self.style.theme_use(theme)
                    break
        
        # Get current user
        self.current_user = self._get_current_user()
        


        # Connection string - Not used directly with pymssql's connect()
        # self.CONNECTION_STRING = f'DRIVER={{SQL Server}};SERVER={self.SERVER};DATABASE={self.DATABASE};Trusted_Connection=yes;'
        
        # Initialize connection_error flag and task manager
        self.connection_error = False
        try:
            self.task_manager = task_manager
        except Exception as e:
            self.connection_error = True
            messagebox.showerror("Connection Error", 
                                f"Could not connect to the SQL Server database.\n\n"
                                f"Error: {str(e)}\n\n"
                                f"The application will run without database connection.")
            logging.error("Connection error at startup", exc_info=True) # Log the exception
            self.task_manager = None
        
        # Initialize deleted tasks stack
        self.deleted_tasks_stack = []
        self.max_undo_stack_size = 10  # Maximum number of undo operations to keep
        
        # Create main frame
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create header
        self.create_header()
        
        # Create task list frame
        self.create_task_list()
        
        # Create control panel
        self.create_control_panel()
        
        # Create status bar
        self.create_status_bar()
        
        # Set initial status
        self.set_status("Loading tasks...")
        
        # Schedule task loading after UI is displayed
        self.root.after(100, self.refresh_task_list)
        
        # Bind the window close event to cleanup method
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # Start the lock cleanup timer
        self._start_lock_cleanup_timer()
    
    def _get_current_user(self):
        """Get the current user's username."""
        try:
            username = os.getenv('USERNAME') or os.getenv('USER') or 'Unknown User'
            logging.info(f"Current user: {username}") # Log the username
            return username
        except:
            return 'Unknown User'
    
    def on_close(self):
        """Handle application close event - clean up resources and close database connections."""
        try:
            # Close database connection if it exists
            if hasattr(self, 'task_manager') and self.task_manager:
                if hasattr(self.task_manager, 'conn') and self.task_manager.conn:
                    try:
                        # Commit any pending changes
                        self.task_manager.conn.commit()
                        # Close the connection
                        self.task_manager.conn.close()
                        logging.info("Database connection closed properly")
                    except Exception as e:
                        logging.error(f"Error closing database connection: {str(e)}", exc_info=True)
        except Exception as e:
            logging.error(f"Error during application shutdown: {str(e)}", exc_info=True)
        finally:
            # Close the application
            self.root.destroy()
    
    def create_header(self):
        """Create the application header."""
        header_frame = ttk.Frame(self.main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Create a left frame for title and storage indicator
        title_frame = ttk.Frame(header_frame)
        title_frame.pack(side=tk.LEFT, padx=5)
        
        title_label = ttk.Label(title_frame, text="Task Manager", style="Title.TLabel")
        title_label.pack(side=tk.TOP, anchor=tk.W)
        
        # Show storage location indicator under the title
        if self.connection_error:
            self.storage_indicator = ttk.Label(title_frame, text="(No Database)", foreground="red")
        else:
            self.storage_indicator = ttk.Label(title_frame, text="(SQL Server)", foreground="green")
        self.storage_indicator.pack(side=tk.TOP, anchor=tk.W)
        
        # Filter controls
        filter_frame = ttk.Frame(header_frame)
        filter_frame.pack(side=tk.RIGHT, padx=5)
        
        # Category filter
        ttk.Label(filter_frame, text="Category:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.category_var = tk.StringVar(value="All")
        self.category_combo = ttk.Combobox(filter_frame, textvariable=self.category_var, width=15, state="readonly")
        self.category_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.update_category_filter()
        self.category_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_task_list())
        
        # Main Staff filter
        ttk.Label(filter_frame, text="Main Staff:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.main_staff_var = tk.StringVar(value="All")
        self.main_staff_combo = ttk.Combobox(filter_frame, textvariable=self.main_staff_var, width=15, values=USERS, state="readonly")
        self.main_staff_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.main_staff_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_task_list())
        
        # User filter
        ttk.Label(filter_frame, text="User:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.user_var = tk.StringVar(value="All")
        self.user_combo = ttk.Combobox(filter_frame, textvariable=self.user_var, width=15, values=USERS, state="readonly")
        self.user_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.user_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_task_list())
        
        # Show completed checkbox - with fixed width to ensure full text is visible
        ttk.Label(filter_frame, text="Status:").pack(side=tk.LEFT, padx=(0, 5))
        
        status_frame = ttk.Frame(filter_frame)
        status_frame.pack(side=tk.LEFT, padx=(0, 10))
        
        self.show_completed_var = tk.BooleanVar(value=False)
        show_completed_check = ttk.Checkbutton(
            status_frame, 
            text="Show Completed Tasks", 
            variable=self.show_completed_var,
            command=self.refresh_task_list,
            width=22  # Increased width to ensure full text is visible including the "s"
        )
        show_completed_check.pack(side=tk.LEFT)
    
    def create_task_list(self):
        """Create the task list with a treeview."""
        # Create frame with scrollbar
        task_frame = ttk.Frame(self.main_frame)
        task_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Create treeview with both vertical and horizontal scrollbars
        columns = (
            "status", "title", "rev", "applied_vessel",
            "priority", "main_staff", "assigned_to",
            "qtd_mhr", "actual_mhr"  # Add new columns here
        )
        
        # Create a frame to hold the treeview and scrollbars
        tree_frame = ttk.Frame(task_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create the treeview
        self.task_tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        
        # Define standard font for all treeview elements - make it bold
        standard_font = ('TkDefaultFont', 9, 'bold')
        
        # Configure the treeview to use the standard font
        self.style.configure("Treeview", font=standard_font)
        self.style.configure("Treeview.Heading", font=('TkDefaultFont', 9, 'bold'))
        
        # Define headings with proper command for toggling sort direction
        self.task_tree.heading("status", text="Status", command=lambda: self.sort_treeview("status", True))
        self.task_tree.heading("title", text="Equipment Name", command=lambda: self.sort_treeview("title", True))
        self.task_tree.heading("rev", text="Rev", command=lambda: self.sort_treeview("rev", True))
        self.task_tree.heading("applied_vessel", text="Applied Vessel", command=lambda: self.sort_treeview("applied_vessel", True))
        self.task_tree.heading("priority", text="Priority", command=lambda: self.sort_treeview("priority", True))
        self.task_tree.heading("main_staff", text="Main Staff", command=lambda: self.sort_treeview("main_staff", True))
        self.task_tree.heading("assigned_to", text="Assigned To", command=lambda: self.sort_treeview("assigned_to", True))
        self.task_tree.heading("qtd_mhr", text="Qtd Mhr", command=lambda: self.sort_treeview("qtd_mhr", True))  # Heading for Qtd Mhr
        self.task_tree.heading("actual_mhr", text="Actual Mhr", command=lambda: self.sort_treeview("actual_mhr", True))  # Heading for Actual Mhr
        
        # Define optimized column widths - REDUCED WIDTHS FOR BETTER VISIBILITY
        self.task_tree.column("status", width=120, minwidth=100, anchor=tk.W)  # Reduced width
        self.task_tree.column("title", width=250, minwidth=200) # Reduced width
        self.task_tree.column("rev", width=40, minwidth=40, anchor=tk.CENTER) # Reduced width
        self.task_tree.column("applied_vessel", width=100, minwidth=100) # Reduced width
        self.task_tree.column("priority", width=70, minwidth=70, anchor=tk.CENTER) # Reduced width
        self.task_tree.column("main_staff", width=80, minwidth=80, anchor=tk.CENTER) # Reduced width
        self.task_tree.column("assigned_to", width=80, minwidth=80, anchor=tk.CENTER) # Reduced width
        self.task_tree.column("qtd_mhr", width=70, minwidth=70, anchor=tk.CENTER)  # Width for Qtd Mhr - slightly reduced
        self.task_tree.column("actual_mhr", width=70, minwidth=70, anchor=tk.CENTER)  # Width for Actual Mhr - slightly reduced
        
        # Initialize sorting variables
        self.sort_column = None
        self.sort_reverse = False
        
        # Add vertical scrollbar
        v_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.task_tree.yview)
        self.task_tree.configure(yscrollcommand=v_scrollbar.set)
        
        # Add horizontal scrollbar
        h_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.task_tree.xview)
        self.task_tree.configure(xscrollcommand=h_scrollbar.set)
        
        # Grid layout for treeview and scrollbars to ensure they're always visible
        self.task_tree.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')
        
        # Configure the grid to expand properly
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        # Create a loading indicator overlay
        self.loading_frame = ttk.Frame(tree_frame, style="TFrame")
        self.loading_label = ttk.Label(
            self.loading_frame, 
            text="Loading tasks...", 
            font=('Helvetica', 12),
            padding=20
        )
        self.loading_label.pack(expand=True, fill=tk.BOTH)
        
        # Bind double-click to view task details
        self.task_tree.bind("<Double-1>", self.view_task_details)
        
        # Bind right-click to show context menu
        self.task_tree.bind("<Button-3>", self.show_context_menu)
        
        # Bind single-click to check for link clicks
        self.task_tree.bind("<ButtonRelease-1>", self.check_link_click)
        
        # Bind mouse motion to change cursor over links
        self.task_tree.bind("<Motion>", self.update_cursor)
        
        # Configure tag for links - use bold font for consistency
        self.task_tree.tag_configure("link", foreground="#0066cc", font=standard_font)
        
        # Configure tag for link hover effect
        self.task_tree.tag_configure("link_hover", foreground="#0099ff", font=standard_font)
        
        # Configure tag for link click effect
        self.task_tree.tag_configure("link_click", foreground="#003366", font=standard_font)
        
        # Configure tag for completed tasks with gray text and background
        self.task_tree.tag_configure("completed", foreground="#9e9e9e", background="#e0e0e0", font=standard_font)  # Gray text and light gray background
    
    def show_loading_indicator(self):
        """Show the loading indicator over the treeview."""
        if hasattr(self, 'loading_frame'):
            self.loading_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
            self.root.update_idletasks()
    
    def hide_loading_indicator(self):
        """Hide the loading indicator."""
        if hasattr(self, 'loading_frame'):
            self.loading_frame.place_forget()
            self.root.update_idletasks()

    def _get_due_date_color(self, due_date_str):
        """Calculate background color based on due date proximity."""
        if not due_date_str:
            return "#ffffff"  # White for no due date

        try:
            due_date = datetime.datetime.strptime(due_date_str, "%Y-%m-%d").date()
            today = datetime.date.today()
            days_until_due = (due_date - today).days

            if days_until_due < 0:
                return "#ffcccc"  # More intense light red for overdue
            elif days_until_due == 0:
                return "#ffd9d9"  # More intense red for due today
            elif days_until_due <= 3:
                return "#ffe6e6"  # Light red for near due
            elif days_until_due <= 7:
                return "#fff2cc"  # More intense light yellow for upcoming
            elif days_until_due <= 14:
                return "#e6ffcc"  # More intense light yellow-green
            else:
                return "#ccffcc"  # More intense light green for far future
        except ValueError:
            return "#ffffff"  # White for invalid date format

    def create_control_panel(self):
        """Create the control panel with buttons."""
        # Create a container with a subtle border
        control_frame = ttk.Frame(self.main_frame, style="TFrame")
        control_frame.pack(fill=tk.X, pady=10)
        
        # Left side buttons
        left_buttons = ttk.Frame(control_frame, style="TFrame")
        left_buttons.pack(side=tk.LEFT)
        
        # Add task button with icon
        add_btn = ttk.Button(
            left_buttons, 
            text=" Add Task", 
            command=self.add_task,
            style="TButton"
        )
        add_btn.pack(side=tk.LEFT, padx=5)
        
        # Edit task button with icon
        edit_btn = ttk.Button(
            left_buttons, 
            text=" Edit Task", 
            command=self.edit_task,
            style="TButton"
        )
        edit_btn.pack(side=tk.LEFT, padx=5)
        
        # Complete task button with icon
        complete_btn = ttk.Button(
            left_buttons, 
            text=" Toggle Status", 
            command=self.complete_task,
            style="Success.TButton"
        )
        complete_btn.pack(side=tk.LEFT, padx=5)
        
        # Delete task button with icon
        delete_btn = ttk.Button(
            left_buttons, 
            text=" Delete Task", 
            command=self.delete_task,
            style="Accent.TButton"
        )
        delete_btn.pack(side=tk.LEFT, padx=5)
        
        # Add Undo button
        self.undo_btn = ttk.Button(
            left_buttons,
            text=" Undo",
            command=self.undo_delete,
            state=tk.DISABLED,  # Initially disabled
            style="TButton"
        )
        self.undo_btn.pack(side=tk.LEFT, padx=5)
        
        # Right side buttons
        right_buttons = ttk.Frame(control_frame, style="TFrame")
        right_buttons.pack(side=tk.RIGHT)
        
        # Refresh button with icon
        refresh_btn = ttk.Button(
            right_buttons, 
            text=" Refresh", 
            command=self.refresh_task_list,
            style="TButton"
        )
        refresh_btn.pack(side=tk.RIGHT, padx=5)
    
    def create_status_bar(self):
        """Create the status bar at the bottom of the main window."""
        status_frame = ttk.Frame(self.root, relief=tk.FLAT, padding="2 2 2 2")
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        separator = ttk.Separator(status_frame, orient=tk.HORIZONTAL)
        separator.pack(fill=tk.X, pady=(0, 2))
        
        # Status bar with task statistics
        self.status_var = tk.StringVar()
        status_bar = ttk.Label(
            status_frame, 
            textvariable=self.status_var, 
            anchor=tk.W,
            style="Status.TLabel"
        )
        status_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Add user info to the right side of the status bar
        username = self.task_manager.get_sql_username()
        if username == "TaskUser1":
            display_name = "Jay"
        elif username == "TaskUser2":
            display_name = "Jude"
        elif username == "TaskUser3":
            display_name = "Jorgen"
        elif username == "TaskUser4":
            display_name = "Earl"
        elif username == "TaskUser5":
            display_name = "Philip"
        elif username == "TaskUser6":
            display_name = "Samuel"
        elif username == "TaskUser7":
            display_name = "Glenn"
        else:
            display_name = username  # Default to username if not in the list

        user_label = ttk.Label(
            status_frame, 
            text=f"User: {display_name}",
            anchor=tk.E,
            style="Status.TLabel"
        )
        user_label.pack(side=tk.RIGHT, padx=10)
        
        # Initialize status
        self.update_status()
    
    def update_status(self):
        """Update the status bar with task statistics."""
        # Only count non-deleted tasks
        active_tasks = [task for task in self.task_manager.tasks if not task.get("deleted", False)]
        total_tasks = len(active_tasks)
        completed = sum(1 for task in active_tasks if task["completed"])
        pending = total_tasks - completed
        
        # Create a more informative status message
        status_text = f"Total Tasks: {total_tasks} | Completed: {completed} | Pending: {pending}"
        
        # Add filter information if filters are active
        filters_active = []
        
        if self.category_var.get() != "All":
            filters_active.append(f"Category: {self.category_var.get()}")
            
        if self.main_staff_var.get() != "All":
            filters_active.append(f"Main Staff: {self.main_staff_var.get()}")
            
        if self.user_var.get() != "All":
            filters_active.append(f"Assigned To: {self.user_var.get()}")
            
        if hasattr(self, 'search_var') and self.search_var.get():
            filters_active.append(f"Search: '{self.search_var.get()}'")
        
        if filters_active:
            status_text += f" | Filters: {', '.join(filters_active)}"
        
        self.status_var.set(status_text)
        
        # Log the updated counts
        logging.info(f"Updated task counts (excluding deleted): Total={total_tasks}, Completed={completed}, Pending={pending}")
    
    def refresh_task_list(self):
        """Refresh the task list in the treeview."""
        # Show loading indicator
        self.set_status("Loading tasks...")
        self.show_loading_indicator()

        # Use threading to perform task refresh in the background
        threading.Thread(target=self._perform_task_refresh_threaded, daemon=True).start()

    def _perform_task_refresh_threaded(self):
        """Threaded method to perform task refresh."""
        try:
            self._perform_task_refresh() # Call the original refresh method
        except Exception as e:
            logging.error("Error during task refresh in thread", exc_info=True)
            # Handle errors that occur during task refresh in the background thread
            self.root.after(0, lambda: messagebox.showerror("Error", f"Error refreshing task list: {str(e)}", parent=self.root)) # Use lambda for deferred execution
        finally:
            # Ensure loading indicator is hidden even if there's an error
            self.root.after(0, self.hide_loading_indicator) # Use root.after to run in main thread
            self.root.after(0, self.update_status) # Update status bar in main thread

    def _perform_task_refresh(self):
        """Perform the actual task refresh after UI has updated."""
        try:
            # First, explicitly reload all tasks from database to get the latest changes
            try:
                self.task_manager.reload_tasks() # Force reload tasks from database
            except Exception as e:
                logging.error(f"Error reloading tasks: {str(e)}", exc_info=True)
                raise Exception(f"Database error: {str(e)}")

            # Clear existing items in the treeview
            for item in self.task_tree.get_children():
                self.task_tree.delete(item)

            # Get filtered tasks
            show_completed = self.show_completed_var.get()
            category = None if self.category_var.get() == "All" else self.category_var.get()
            main_staff = None if self.main_staff_var.get() == "All" else self.main_staff_var.get()
            assigned_to = None if self.user_var.get() == "All" else self.user_var.get()

            # Get search term if available
            search_term = ""
            if hasattr(self, 'search_var'):
                search_term = self.search_var.get().lower().strip()

            # Get filtered tasks
            filtered_tasks = self.task_manager.get_filtered_tasks(
                show_completed=show_completed,
                category=category,
                main_staff=main_staff,
                assigned_to=assigned_to
            )

            # Apply search filter if needed
            if search_term:
                filtered_tasks = [
                    task for task in filtered_tasks if (
                        search_term in task["title"].lower() or
                        search_term in (task.get("description", "")).lower() or
                        search_term in (task.get("applied_vessel", "")).lower() or
                        search_term in (task.get("category", "")).lower() or
                        search_term in (task.get("main_staff", "")).lower() or
                        search_term in (task.get("assigned_to", "")).lower()
                    )
                ]

            # Apply sorting if a sort column is set
            if self.sort_column:
                self.apply_sorting(filtered_tasks)

            # Track the maximum status text length to adjust column width
            max_status_length = 0

            # Define a standard font for all task items - make it bold
            standard_font = ('TkDefaultFont', 9, 'bold')

            # Add tasks to treeview with improved visual indicators
            for task in filtered_tasks:
                # Create status with time remaining - using standardized format for better alignment
                if task["completed"]:
                    status = "✓ Completed"
                    status_tags = ["completed"]
                else:
                    # Add due date information for pending tasks
                    if task["due_date"]:
                        try:
                            due_date = datetime.datetime.strptime(task["due_date"], "%Y-%m-%d").date()
                            today = datetime.date.today()
                            days_until_due = (due_date - today).days
                            
                            if days_until_due < 0:
                                # Use consistent format for overdue
                                status = f"! {abs(days_until_due)} Days Overdue"
                                status_tags = ["overdue"]
                            elif days_until_due == 0:
                                status = "! Due Today"
                                status_tags = ["due_today"]
                            elif days_until_due == 1:
                                status = "! Due Tomorrow"
                                status_tags = ["pending_soon"]
                            else:
                                # Use consistent format for days left
                                status = f"• {days_until_due} Days Left"
                                status_tags = ["pending"]
                                if days_until_due <= 3:
                                    status_tags = ["pending_soon"]
                        except ValueError:
                            status = "• No Due Date"
                            status_tags = ["pending"]
                    else:
                        status = "• No Due Date"
                        status_tags = ["pending"]
                
                # Update max status length for column width adjustment
                max_status_length = max(max_status_length, len(status))
                
                # Format due date
                due_date = task["due_date"] if task["due_date"] else "No due date"
                
                # Format staff assignments
                main_staff = task.get("main_staff", "Unassigned")
                assigned_to = task.get("assigned_to", "Unassigned")
                
                # Get values for new columns (with defaults if not present)
                applied_vessel = task.get("applied_vessel", "")
                rev = task.get("rev", "")
                
                # Set row tags for styling based on priority and completion
                if task["completed"]:
                    # For completed tasks, only use the completed tag to ensure it overrides all other styles
                    tags = ["completed"]
                else:
                    # For pending tasks, use priority and status tags
                    tags = [task["priority"]] + status_tags
                
                # Format priority with emoji (without link indicator)
                priority_display = {
                    "high": "🔴 High",
                    "medium": "🔵 Medium",
                    "low": "🟢 Low"
                }.get(task["priority"].lower(), task["priority"].capitalize())
                
                # Create link display
                link_display = "🔗" if task.get("link") else ""
                
                # Create a unique tag for this task's due date color
                due_date_tag = f"due_date_{task['id']}"
                tags.append(due_date_tag)
                
                # Configure the tag with the appropriate background color
                bg_color = self._get_due_date_color(task["due_date"])
                self.task_tree.tag_configure(due_date_tag, background=bg_color)
                
                # For completed tasks, make sure "completed" is the first tag to ensure it has priority
                if task["completed"]:
                    # Remove any status tags that might override the completed style
                    tags = [tag for tag in tags if tag not in ["overdue", "due_today", "pending_soon", "pending"]]
                    # Put completed tag first to ensure it has priority
                    tags.insert(0, "completed")
                
                # Add link tag if there's a link
                if task.get("link"):
                    tags.append("link")
                
                self.task_tree.insert(
                    "", tk.END, 
                    values=(
                        status,
                        task["title"],
                        task.get("rev", ""),
                        task.get("applied_vessel", ""),
                        priority_display,
                        task.get("main_staff", ""),
                        task.get("assigned_to", ""),
                        task.get("qtd_mhr", ""),  # Qtd Mhr value
                        task.get("actual_mhr", "") # Actual Mhr value
                    ),
                    tags=tags,
                    iid=str(task["id"])  # Use task ID as the item ID for direct lookup
                )
            
            # Adjust the status column width based on content
            # Calculate pixel width: approximate 8 pixels per character plus some padding
            if max_status_length > 0:
                # Use a multiplier for average character width (depends on font)
                char_width = 8  # Approximate width in pixels per character
                padding = 20    # Extra padding for the column
                new_width = max(150, min(300, max_status_length * char_width + padding))
                self.task_tree.column("status", width=new_width, minwidth=150)
            
            # Configure tag colors for status indicators with consistent bold font
            self.task_tree.tag_configure("completed", foreground="#9e9e9e", font=standard_font)  # Gray text
            self.task_tree.tag_configure("overdue", foreground="#d32f2f", font=standard_font)  # Red text
            self.task_tree.tag_configure("due_today", foreground="#f57c00", font=standard_font)  # Orange text
            self.task_tree.tag_configure("pending_soon", foreground="#7b1fa2", font=standard_font)  # Purple text
            self.task_tree.tag_configure("pending", foreground="#1976d2", font=standard_font)  # Blue text
            
            # Configure priority tags with consistent bold font
            self.task_tree.tag_configure("high", font=standard_font)
            self.task_tree.tag_configure("medium", font=standard_font)
            self.task_tree.tag_configure("low", font=standard_font)
            
            # Configure link tag with consistent bold font but still make it stand out
            self.task_tree.tag_configure("link", foreground="#0066cc", font=standard_font)
            self.task_tree.tag_configure("link_hover", foreground="#0099ff", font=standard_font)
            self.task_tree.tag_configure("link_click", foreground="#003366", font=standard_font)
            
            # Show "no results" message if no tasks match the filters
            if not filtered_tasks:
                self.task_tree.insert("", tk.END, values=("No tasks match your filters", "", "", "", "", "", "", "", ""), tags=["no_results"])
                self.task_tree.tag_configure("no_results", foreground="#9e9e9e", font=standard_font)
            
        except Exception as e:
            logging.error("Error during _perform_task_refresh", exc_info=True)
            # Re-raise to be caught in _perform_task_refresh_threaded - No, handle here and show error
            self.root.after(0, lambda: messagebox.showerror("Error", "Error refreshing task list. See log for details.", parent=self.root)) # Show error in main thread
        finally:
            # Hide loading indicator and update status are now handled in _perform_task_refresh_threaded
            pass # No changes needed here
    
    def update_category_filter(self):
        """Update the category filter dropdown with available categories."""
        categories = set(task["category"] for task in self.task_manager.tasks)
        categories = sorted(list(categories))
        
        # Add "All" option
        categories = ["All"] + categories
        
        # Update combobox values
        current = self.category_var.get()
        self.category_combo["values"] = categories
        
        # If current category is not in the list, reset to "All"
        if current not in categories:
            self.category_var.set("All")
    
    def get_selected_task_id(self):
        """Get the ID of the selected task (single selection)."""
        selection = self.task_tree.selection()
        if not selection:
            messagebox.showinfo("Information", "Please select a task first.")
            return None
        
        # Get the task ID directly from the item ID
        try:
            task_id = int(selection[0])
            return task_id
        except (ValueError, TypeError):
            # Fallback to the old method if the item ID is not a valid task ID
            item = selection[0]
            values = self.task_tree.item(item, "values")
            
            # Find the task with matching values
            for task in self.task_manager.tasks:
                # Match by title, rev, and applied_vessel
                if (task["title"] == values[1] and
                    task.get("rev", "") == values[2] and
                    task.get("applied_vessel", "") == values[3]):
                    return task["id"]
            
            messagebox.showerror("Error", "Could not identify the selected task.")
            return None
    
    def get_selected_task_ids(self):
        """Get the IDs of the selected tasks."""
        selection = self.task_tree.selection()
        if not selection:
            messagebox.showinfo("Information", "Please select at least one task.")
            return None
        
        # Get the IDs directly from the item IDs
        task_ids = []
        for item in selection:
            try:
                task_id = int(item)
                task_ids.append(task_id)
            except (ValueError, TypeError):
                # Fallback to the old method if the item ID is not a valid task ID
                values = self.task_tree.item(item, "values")
                
                # Find the task with matching values
                for task in self.task_manager.tasks:
                    # Match by title, rev, and applied_vessel
                    if (task["title"] == values[1] and
                        task.get("rev", "") == values[2] and
                        task.get("applied_vessel", "") == values[3]):
                        task_ids.append(task["id"])
                        break
        
        if not task_ids:
            messagebox.showerror("Error", "Could not identify the selected tasks.")
            return None
            
        return task_ids
    
    def add_task(self):
        """Open dialog to add a new task."""
        dialog = TaskDialog(self.root, "Add Task")
        if dialog.result:
            # Add the task
            self.task_manager.add_task(**dialog.result)
            self.refresh_task_list()
            self.set_status(f"Task '{dialog.result['title']}' added successfully.")
    
    def edit_task(self):
        """Edit task with locking mechanism to prevent concurrent edits."""
        task_id = self.get_selected_task_id()
        if not task_id:
            messagebox.showinfo("No Task Selected", "Please select a task to edit.")
            return
        
        current_user = self._get_current_user()
        
        # Try to lock the task for editing
        if not self.task_manager.lock_task_for_editing(task_id, current_user):
            # Get lock information using task_manager's cursor, not self.cursor
            try:
                self.task_manager.cursor.execute("SELECT locked_by FROM Tasks WHERE id = %s", (task_id,))
                result = self.task_manager.cursor.fetchone()
                locked_by = result[0] if result else "another user"
                
                messagebox.showwarning(
                    "Task Locked", 
                    f"This task is currently being edited by {locked_by}.\n"
                    "Please try again later."
                )
            except Exception as e:
                logging.error(f"Error checking lock status: {str(e)}", exc_info=True)
                messagebox.showwarning(
                    "Task Locked", 
                    "This task is currently being edited by another user.\n"
                    "Please try again later."
                )
            return
        
        # Task is now locked, proceed with edit
        task = self.task_manager._find_task_by_id(task_id)
        if not task:
            self.task_manager.release_task_lock(task_id, current_user)
            messagebox.showerror("Error", "Task not found.")
            return
        
        # Store the current version for concurrency control
        current_version = task.get('version', 1)
        
        try:
            dialog = TaskDialog(self.master, "Edit Task", task)
            if dialog.result:
                # Add expected version to the task data for concurrency checking
                dialog.result['expected_version'] = current_version
                
                # Try to update with concurrency control
                success = self.task_manager.update_task(task_id, **dialog.result)
                
                if success:
                    self.refresh_task_list()
                    self.set_status(f"Task '{dialog.result['title']}' updated successfully.")
                else:
                    # Handle concurrency conflict
                    conflict_response = messagebox.askyesno(
                        "Update Conflict", 
                        "This task has been modified by another user since you opened it.\n\n"
                        "Would you like to reload the task and try again?",
                        icon="warning"
                    )
                    
                    if conflict_response:
                        # Refresh and try again
                        self.refresh_task_list()
                        # After refresh, try to select the same task again
                        self.tree.selection_set(task_id)
                        self.edit_task()
                    else:
                        self.set_status("Task update cancelled due to conflict with another user.")
        finally:
            # Always release the lock when done
            self.task_manager.release_task_lock(task_id, current_user)
    
    def complete_task(self):
        """Mark the selected task(s) as completed or pending."""
        task_ids = self.get_selected_task_ids()
        if not task_ids:
            return
        
        # Check if all selected tasks have the same completion status
        tasks = [self.task_manager._find_task_by_id(task_id) for task_id in task_ids]
        all_completed = all(task["completed"] for task in tasks if task)
        all_pending = all(not task["completed"] for task in tasks if task)
        
        count = len(task_ids)
        
        # If all tasks have the same status, offer to toggle them
        if all_completed:
            message = f"Mark {count} selected task{'s' if count > 1 else ''} as pending?"
            if messagebox.askyesno("Confirm", message):
                self.batch_update_completion_status(task_ids, completed=False)
                self.set_status(f"{count} task{'s' if count > 1 else ''} marked as pending.")
        elif all_pending:
            message = f"Mark {count} selected task{'s' if count > 1 else ''} as completed?"
            if messagebox.askyesno("Confirm", message):
                self.batch_update_completion_status(task_ids, completed=True)
                self.set_status(f"{count} task{'s' if count > 1 else ''} marked as completed.")
        else:
            # Mixed status - ask what to do
            options = ["Mark All as Completed", "Mark All as Pending", "Cancel"]
            choice = messagebox.askquestion("Mixed Status", 
                                          f"The {count} selected tasks have mixed completion statuses.\n\n"
                                          "What would you like to do?",
                                          type=messagebox.YESNOCANCEL,
                                          icon=messagebox.QUESTION)
            
            if choice == "yes":  # Yes corresponds to the first option
                self.batch_update_completion_status(task_ids, completed=True)
                self.set_status(f"{count} task{'s' if count > 1 else ''} marked as completed.")
            elif choice == "no":  # No corresponds to the second option
                self.batch_update_completion_status(task_ids, completed=False)
                self.set_status(f"{count} task{'s' if count > 1 else ''} marked as pending.")
        
        self.refresh_task_list()
    
    def batch_update_completion_status(self, task_ids, completed):
        """Update completion status for multiple tasks at once."""
        for task_id in task_ids:
            self.task_manager.update_task(task_id, completed=completed)
        
        # No need to call refresh_task_list() here as it will be called by the calling method
    
    def delete_task(self):
        """Delete the selected task(s)."""
        task_ids = self.get_selected_task_ids()
        if not task_ids:
            return
        
        count = len(task_ids)
        confirm_message = f"Are you sure you want to delete {count} task{'s' if count > 1 else ''}?\n"
        
        if messagebox.askyesno("Confirm Deletion", confirm_message, icon=messagebox.WARNING):
            try:
                # Get the tasks to be deleted before they're removed from the database
                tasks_to_delete = [self.task_manager._find_task_by_id(task_id) for task_id in task_ids]
                
                # Remove None values from the list (if any)
                tasks_to_delete = [task for task in tasks_to_delete if task is not None]
                
                if not tasks_to_delete:
                    messagebox.showinfo("Information", "No tasks were deleted.")
                    return
                
                # Store the deleted tasks with their IDs and deletion timestamp
                deletion_record = {
                    'timestamp': datetime.datetime.now(),
                    'tasks': tasks_to_delete,
                    'task_ids': task_ids
                }
                
                # Add to the stack and maintain max size
                self.deleted_tasks_stack.append(deletion_record)
                if len(self.deleted_tasks_stack) > self.max_undo_stack_size:
                    self.deleted_tasks_stack.pop(0)
                
                # We'll now use the improved batch_delete_tasks which handles its own transaction
                successful_deletions, failed_deletions = self.task_manager.batch_delete_tasks(task_ids)
                
                if failed_deletions:
                    # The batch_delete_tasks method already handled the rollback
                    # Remove the deletion record
                    self.deleted_tasks_stack.pop()
                    messagebox.showerror("Error", 
                        f"Could not delete {len(failed_deletions)} task{'s' if len(failed_deletions) > 1 else ''}.")
                    return
                
                # Enable the Undo button
                self.undo_btn.config(state=tk.NORMAL)
                self.set_status(f"{len(successful_deletions)} task{'s' if len(successful_deletions) > 1 else ''} deleted successfully.")
                
                # Remove the deleted task IDs from the treeview directly
                for task_id in successful_deletions:
                    try:
                        self.task_tree.delete(str(task_id))
                    except:
                        pass  # Item might already be gone
                
                # Force refresh the task list to ensure consistency
                self.refresh_task_list()
                
            except Exception as e:
                logging.error("Error during task deletion", exc_info=True)
                messagebox.showerror("Error", f"An error occurred while deleting tasks: {str(e)}")
    
    def view_task_details(self, event):
        """View details of the selected task."""
        # If called from an event, get the task ID from the clicked item
        if event:
            item = self.task_tree.identify_row(event.y)
            if not item:
                return
            self.task_tree.selection_set(item)
        
        task_id = self.get_selected_task_id()
        if not task_id:
            return
        
        # Get the task - reload tasks first to ensure we have the latest data
        self.task_manager.reload_tasks()
        task = self.task_manager._find_task_by_id(task_id)
        if not task:
            messagebox.showerror("Error", f"Task with ID {task_id} not found.")
            return
        
        # Format dates for display
        start_date = task.get('date_started', 'Not specified')
        request_date = task.get('requested_date', 'Not specified')
        due_date = task.get('due_date', 'No due date')
        created_date = task.get('created_date', 'Unknown')
        last_modified = task.get('last_modified', 'Not modified')
        modified_by = task.get('modified_by', 'Not modified')
        
        # Show task details
        details = (
            f"Equipment Name: {task['title']}\n\n"
            f"Rev: {task.get('rev', 'Not specified')}\n\n"
            f"Description: {task['description'] or 'No description'}\n\n"
            f"Start: {start_date}\n\n"
            f"Request: {request_date}\n\n"
            f"Due: {due_date}\n\n"
            f"Category: {task['category']}\n\n"
            f"Created date: {created_date}\n\n"
            f"Last Modified: {last_modified}\n\n"
            f"Modified by: {modified_by}"
        )
        
        messagebox.showinfo(f"Task Details", details)
    
    def show_context_menu(self, event):
        """Show the context menu for tasks."""
        # Get the item under cursor
        item = self.task_tree.identify_row(event.y)
        if not item:
            return
        
        # Select the item if it's not already selected
        if item not in self.task_tree.selection():
            self.task_tree.selection_set(item)
        
        # Get the number of selected items
        selection_count = len(self.task_tree.selection())
        
        # Create context menu
        context_menu = tk.Menu(self.root, tearoff=0)
        
        # Task actions submenu
        task_menu = tk.Menu(context_menu, tearoff=0)
        if selection_count == 1:
            task_menu.add_command(label="View Details", command=lambda: self.view_task_details(None))
            task_menu.add_command(label="Edit Task", command=self.edit_task)
            task_menu.add_command(label="Toggle Complete/Pending", command=self.complete_task)
        else:
            task_menu.add_command(label=f"Toggle {selection_count} Tasks", command=self.complete_task)
        
        context_menu.add_cascade(label="Task Actions", menu=task_menu)
        
        # Add Assign To submenu
        assign_menu = tk.Menu(context_menu, tearoff=0)
        for staff in USERS[1:]:  # Skip "All" from the USERS list
            assign_menu.add_command(
                label=staff,
                command=lambda s=staff: self.assign_task_to(s)
            )
        context_menu.add_cascade(label="Assign To", menu=assign_menu)
        
        # Copy options submenu (only for single selection)
        if selection_count == 1:
            try:
                task_id = int(item)
                task = self.task_manager._find_task_by_id(task_id)
                if task:
                    # Create Copy submenu
                    copy_menu = tk.Menu(context_menu, tearoff=0)
                    has_copy_items = False
                    
                    # Add copy options for hidden fields
                    if task.get("drawing_no"):
                        copy_menu.add_command(
                            label="Copy Drawing No.", 
                            command=lambda: self.copy_to_clipboard(task.get("drawing_no", ""))
                        )
                        has_copy_items = True
                    
                    if task.get("request_no"):
                        copy_menu.add_command(
                            label="Copy Request No.", 
                            command=lambda: self.copy_to_clipboard(task.get("request_no", ""))
                        )
                        has_copy_items = True
                    
                    if task.get("link"):
                        copy_menu.add_command(
                            label="Copy Folder Link", 
                            command=lambda: self.copy_to_clipboard(task.get("link", ""))
                        )
                        has_copy_items = True
                    
                    if task.get("sdb_link"):
                        copy_menu.add_command(
                            label="Copy SDB Link", 
                            command=lambda: self.copy_to_clipboard(task.get("sdb_link", ""))
                        )
                        has_copy_items = True
                    
                    # Add dates submenu if any dates are present
                    dates = {
                        "Requested Date": task.get("requested_date"),
                        "Date Started": task.get("date_started"),
                        "Due Date": task.get("due_date")
                    }
                    
                    if any(dates.values()):
                        dates_menu = tk.Menu(copy_menu, tearoff=0)
                        for label, date in dates.items():
                            if date:
                                dates_menu.add_command(
                                    label=f"Copy {label}", 
                                    command=lambda d=date: self.copy_to_clipboard(d)
                                )
                        copy_menu.add_cascade(label="Copy Dates", menu=dates_menu)
                        has_copy_items = True
                    
                    # Only add the Copy menu if there are items in it
                    if has_copy_items:
                        context_menu.add_cascade(label="Copy", menu=copy_menu)
                    
                    # Create Open submenu
                    open_menu = tk.Menu(context_menu, tearoff=0)
                    has_open_items = False
                    
                    if task.get("link"):
                        open_menu.add_command(
                            label="Open Folder Link", 
                            command=lambda: self.open_link(task.get("link", ""))
                        )
                        has_open_items = True
                    
                    if task.get("sdb_link"):
                        open_menu.add_command(
                            label="Open SDB Link", 
                            command=lambda: self.open_link(task.get("sdb_link", ""))
                        )
                        has_open_items = True
                    
                    # Only add the Open menu if there are items in it
                    if has_open_items:
                        context_menu.add_cascade(label="Open", menu=open_menu)
                    
            except (ValueError, TypeError):
                # If we can't get the task ID directly, try to find the task by values
                values = self.task_tree.item(item, "values")
                if len(values) >= 4:  # Make sure we have enough values
                    for task in self.task_manager.tasks:
                        # Match by title, rev, and applied_vessel
                        if (task["title"] == values[1] and
                            task.get("rev", "") == values[2] and
                            task.get("applied_vessel", "") == values[3]):
                            
                            # Create Copy submenu
                            copy_menu = tk.Menu(context_menu, tearoff=0)
                            has_copy_items = False
                            
                            # Add copy options for hidden fields
                            if task.get("drawing_no"):
                                copy_menu.add_command(
                                    label="Copy Drawing No.", 
                                    command=lambda: self.copy_to_clipboard(task.get("drawing_no", ""))
                                )
                                has_copy_items = True
                            
                            if task.get("request_no"):
                                copy_menu.add_command(
                                    label="Copy Request No.", 
                                    command=lambda: self.copy_to_clipboard(task.get("request_no", ""))
                                )
                                has_copy_items = True
                            
                            if task.get("link"):
                                copy_menu.add_command(
                                    label="Copy Folder Link", 
                                    command=lambda: self.copy_to_clipboard(task.get("link", ""))
                                )
                                has_copy_items = True
                            
                            if task.get("sdb_link"):
                                copy_menu.add_command(
                                    label="Copy SDB Link", 
                                    command=lambda: self.copy_to_clipboard(task.get("sdb_link", ""))
                                )
                                has_copy_items = True
                            
                            # Add dates submenu if any dates are present
                            dates = {
                                "Requested Date": task.get("requested_date"),
                                "Date Started": task.get("date_started"),
                                "Due Date": task.get("due_date")
                            }
                            
                            if any(dates.values()):
                                dates_menu = tk.Menu(copy_menu, tearoff=0)
                                for label, date in dates.items():
                                    if date:
                                        dates_menu.add_command(
                                            label=f"Copy {label}", 
                                            command=lambda d=date: self.copy_to_clipboard(d)
                                        )
                                copy_menu.add_cascade(label="Copy Dates", menu=dates_menu)
                                has_copy_items = True
                            
                            # Only add the Copy menu if there are items in it
                            if has_copy_items:
                                context_menu.add_cascade(label="Copy", menu=copy_menu)
                            
                            # Create Open submenu
                            open_menu = tk.Menu(context_menu, tearoff=0)
                            has_open_items = False
                            
                            if task.get("link"):
                                open_menu.add_command(
                                    label="Open Folder Link", 
                                    command=lambda: self.open_link(task.get("link", ""))
                                )
                                has_open_items = True
                            
                            if task.get("sdb_link"):
                                open_menu.add_command(
                                    label="Open SDB Link", 
                                    command=lambda: self.open_link(task.get("sdb_link", ""))
                                )
                                has_open_items = True
                            
                            # Only add the Open menu if there are items in it
                            if has_open_items:
                                context_menu.add_cascade(label="Open", menu=open_menu)
                            
                            break
        
        context_menu.add_separator()
        
        # Delete option shows count of items to be deleted
        if selection_count == 1:
            context_menu.add_command(label="Delete Task", command=self.delete_task)
        else:
            context_menu.add_command(label=f"Delete {selection_count} Tasks", command=self.delete_task)
        
        # Add a 'View Deleted Tasks' option
        context_menu.add_separator()
        context_menu.add_command(label="View Deleted Tasks", command=self.view_deleted_tasks)
        
        # Display the context menu
        context_menu.tk_popup(event.x_root, event.y_root)
    
    def check_link_click(self, event):
        """Check if a link cell was clicked and open the link if so."""
        region = self.task_tree.identify_region(event.x, event.y)
        if region == "cell":
            # Get the item and column that was clicked
            item = self.task_tree.identify_row(event.y)
            column = self.task_tree.identify_column(event.x)
            
            # Check if the link column was clicked (column #8)
            if column == "#8":  # Link column
                # Get the values from the display
                values = self.task_tree.item(item, "values")
                link_display = values[7] if len(values) > 7 else ""
                
                # Only proceed if it contains a link indicator
                if link_display == "🔗":
                    # Get the actual task to find the real link
                    try:
                        task_id = int(item)
                        task = self.task_manager._find_task_by_id(task_id)
                        if task and task.get("link"):
                            # Apply click effect
                            tags = list(self.task_tree.item(item, "tags"))
                            if "link" in tags:
                                tags.remove("link")
                            if "link_hover" in tags:
                                tags.remove("link_hover")
                            tags.append("link_click")
                            self.task_tree.item(item, tags=tags)
                            
                            # Schedule to revert the effect after a short delay
                            self.root.after(150, lambda i=item: self.revert_link_click_effect(i))
                            
                            # Open the link
                            self.open_link(task["link"])
                    except (ValueError, TypeError):
                        # If we can't get the task ID directly, try to find the task
                        for task in self.task_manager.tasks:
                            # Check if this is the right task by matching other values
                            if (task["title"] == values[1] and
                                task.get("rev", "") == values[2] and
                                task.get("applied_vessel", "") == values[3]):
                                if task.get("link"):
                                    # Open the link
                                    self.open_link(task["link"])
                                break
    
    def revert_link_click_effect(self, item):
        """Revert the link click effect after a short delay."""
        if item in self.task_tree.get_children():
            tags = list(self.task_tree.item(item, "tags"))
            if "link_click" in tags:
                tags.remove("link_click")
            if "link" not in tags:
                tags.append("link")
            self.task_tree.item(item, tags=tags)
    
    def open_link(self, link):
        """Open a link which can be a URL, file path, or network path."""
        
        if not link:
            messagebox.showerror("Error", "No link provided.")
            return
            
        try:
            # Check if it's a network path or local file path
            if link.startswith("\\\\") or os.path.exists(link) or link.startswith(("C:", "D:", "E:", "F:", "\\")):
                # It's a file path - use the appropriate command based on OS
                if os.name == 'nt':  # Windows
                    # For network paths, ensure they're properly formatted
                    if link.startswith("//"):
                        # Convert forward slashes to backslashes for Windows
                        link = link.replace("/", "\\")
                    
                    # Check if it's a directory
                    if os.path.isdir(link) or (link.startswith("\\\\") and not os.path.splitext(link)[1]):
                        # Open directory in Explorer
                        subprocess.run(['explorer', link], shell=True)
                    else:
                        # Open file with default application
                        os.startfile(link)
                elif os.name == 'posix':  # macOS and Linux
                    subprocess.run(['xdg-open', link] if os.name == 'posix' else ['open', link])
            else:
                # It's a web URL - add http:// if needed
                if not link.startswith(("http://", "https://", "www.")):
                    link = f"http://{link}"
                webbrowser.open(link)
                
            self.set_status(f"Opening: {link}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Could not open link: {link}\n\nError: {str(e)}")
            # Log more detailed error information
            print(f"Error opening link: {link}")
            print(f"Error details: {str(e)}")
            print(f"Error type: {type(e)}")
    
    def set_status(self, message):
        """Set a temporary status message."""
        self.status_var.set(message)
        # Schedule to reset after 5 seconds
        self.root.after(5000, self.update_status)
    
    def update_cursor(self, event):
        """Update cursor based on whether it's over a link."""
        region = self.task_tree.identify_region(event.x, event.y)
        if region == "cell":
            column = self.task_tree.identify_column(event.x)
            item = self.task_tree.identify_row(event.y)
            
            # Reset all link items to normal link style
            for all_item in self.task_tree.get_children():
                tags = list(self.task_tree.item(all_item, "tags"))
                if "link_hover" in tags:
                    tags.remove("link_hover")
                    if "link" not in tags:
                        tags.append("link")
                    self.task_tree.item(all_item, tags=tags)
            
            # Check if over the link column (column #8)
            if column == "#8" and item:
                # Get the values from the display
                values = self.task_tree.item(item, "values")
                link_display = values[7] if len(values) > 7 else ""
                
                # Only change cursor if it contains a link indicator
                if link_display == "🔗":
                    # Get the actual task to confirm it has a link
                    try:
                        task_id = int(item)
                        task = self.task_manager._find_task_by_id(task_id)
                        if task and task.get("link"):
                            self.task_tree.config(cursor="hand2")
                            
                            # Apply hover effect to this link
                            tags = list(self.task_tree.item(item, "tags"))
                            if "link" in tags:
                                tags.remove("link")
                            if "link_hover" not in tags:
                                tags.append("link_hover")
                            self.task_tree.item(item, tags=tags)
                            return
                    except (ValueError, TypeError):
                        # Try to find the task by values
                        for task in self.task_manager.tasks:
                            if (task["title"] == values[1] and
                                task.get("rev", "") == values[2] and
                                task.get("applied_vessel", "") == values[3]):
                                if task.get("link"):
                                    self.task_tree.config(cursor="hand2")
                                    
                                    # Apply hover effect to this link
                                    tags = list(self.task_tree.item(item, "tags"))
                                    if "link" in tags:
                                        tags.remove("link")
                                    if "link_hover" not in tags:
                                        tags.append("link_hover")
                                    self.task_tree.item(item, tags=tags)
                                    return
                                break
            
            # Reset cursor if not over a link
            self.task_tree.config(cursor="")
    
    def copy_to_clipboard(self, text):
        """Copy text to clipboard."""
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.set_status(f"Copied to clipboard: {text}")
    
    def sort_treeview(self, column, reset=True):
        """
        Sort treeview by a column.
        
        Args:
            column: Column to sort by
            reset: Whether to reset the sort direction if clicking the same column
        """
        # Update sort indicators in column headings
        for col in self.task_tree["columns"]:
            # Remove any existing sort indicators
            heading_text = self.task_tree.heading(col)["text"]
            heading_text = heading_text.replace(" ▲", "").replace(" ▼", "")
            self.task_tree.heading(col, text=heading_text)
        
        # Determine sort direction
        if self.sort_column == column and reset:
            # Toggle sort direction if clicking the same column
            self.sort_reverse = not self.sort_reverse
        else:
            # Default to ascending for a new column
            self.sort_column = column
            self.sort_reverse = False
        
        # Add sort indicator to column heading
        heading_text = self.task_tree.heading(column)["text"]
        indicator = " ▲" if not self.sort_reverse else " ▼"
        self.task_tree.heading(column, text=heading_text + indicator)
        
        # Refresh the task list to apply the sorting
        self.refresh_task_list()
    
    def apply_sorting(self, tasks):
        """
        Apply sorting to the task list.
        
        Args:
            tasks: List of tasks to sort
        """
        # Define custom key functions for special columns
        def priority_key(task):
            # Sort by priority level (high, medium, low)
            priority_order = {"high": 0, "medium": 1, "low": 2}
            return priority_order.get(task.get("priority", "").lower(), 3)
        
        def status_key(task):
            # Sort by completion status and due date
            if task.get("completed", False):
                return (2, "")  # Completed tasks at the bottom
            
            if not task.get("due_date"):
                return (1, "")  # Tasks without due date in the middle
            
            # Tasks with due date sorted by days until due
            try:
                due_date = datetime.datetime.strptime(task["due_date"], "%Y-%m-%d").date()
                today = datetime.date.today()
                days_until_due = (due_date - today).days
                
                # Return a tuple for proper sorting
                if days_until_due < 0:
                    return (0, f"{abs(days_until_due):04d}")  # Overdue tasks first
                else:
                    return (0, f"{days_until_due+10000:04d}")  # Then by days until due
            except ValueError:
                return (1, "")  # Invalid date format in the middle
        
        # Sort the tasks based on the selected column
        if self.sort_column == "status":
            tasks.sort(key=status_key, reverse=self.sort_reverse)
        elif self.sort_column == "priority":
            tasks.sort(key=priority_key, reverse=self.sort_reverse)
        elif self.sort_column == "rev":
            # Special handling for revision numbers (numeric sorting when possible)
            def rev_key(task):
                rev = task.get("rev", "")
                # Try to extract numeric part for numeric sorting
                try:
                    # If it's a pure number, convert to int for proper sorting
                    if rev.isdigit():
                        return int(rev)
                    # If it has a format like "Rev 1" or "R2", extract the number
                    import re
                    match = re.search(r'(\d+)', rev)
                    if match:
                        return int(match.group(1))
                except:
                    pass
                # Fall back to string sorting
                return rev.lower()
            
            tasks.sort(key=rev_key, reverse=self.sort_reverse)
        else:
            # For other columns, sort by the column value
            # Use a case-insensitive sort for text columns
            tasks.sort(
                key=lambda task: str(task.get(self.sort_column, "")).lower(),
                reverse=self.sort_reverse
            )

    def assign_task_to(self, staff_name):
        """Assign the selected task(s) to a staff member."""
        task_ids = self.get_selected_task_ids()
        if not task_ids:
            return
            
        count = len(task_ids)
        # Update each selected task
        for task_id in task_ids:
            self.task_manager.update_task(task_id, assigned_to=staff_name)
        
        # Refresh the task list and show confirmation
        self.refresh_task_list()
        self.set_status(f"{count} task{'s' if count > 1 else ''} assigned to {staff_name}")
        
        # Disable the Undo button after any other action
        if hasattr(self, 'undo_btn'):
            self.undo_btn.config(state=tk.DISABLED)

    def undo_delete(self):
        """Recover the most recently deleted batch of tasks within a time limit."""
        if not self.deleted_tasks_stack:
            messagebox.showinfo("Undo Delete", "No deleted tasks to recover")
            return
            
        # Get the most recent deletion record
        deletion_record = self.deleted_tasks_stack.pop()
        
        # Check if the deletion is within the time limit (5 minutes)
        time_limit = datetime.datetime.now() - datetime.timedelta(minutes=5)
        if deletion_record['timestamp'] < time_limit:
            messagebox.showinfo("Undo Delete", "Deletion is too old to recover (over 5 minutes)")
            return
            
        # Recover all tasks in this batch
        recovered_count = 0
        for task_id in deletion_record['task_ids']:
            if self.task_manager.recover_task(task_id):
                recovered_count += 1
        
        # Refresh the task list
        self.refresh_task_list()
        
        # Update status message
        if recovered_count > 0:
            self.set_status(f"Recovered {recovered_count} task{'s' if recovered_count > 1 else ''}")
        else:
            messagebox.showerror("Error", "Failed to recover tasks")
        
        # Disable the undo button if there are no more items in the stack
        if not self.deleted_tasks_stack:
            self.undo_btn.config(state=tk.DISABLED)
    
    def position_window_center(self):
        """Positions the application window in the center of the screen."""
        self.root.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = self.root.winfo_width()
        window_height = self.root.winfo_height()
        x = int((screen_width / 2) - (window_width / 2))
        y = int((screen_height / 2) - (window_height / 2))
        self.root.geometry(f"+{x}+{y}")

    def bind_mousewheel_to_treeview(self):
        """Binds mousewheel scrolling to the task treeview for different platforms."""
        if platform.system() == 'Windows':
            self.task_tree.bind("<MouseWheel>", self.on_mousewheel_windows)
        elif platform.system() == 'Darwin':  # macOS
            self.task_tree.bind("<MouseWheel>", self.on_mousewheel_macos)
        elif platform.system() == 'Linux':
            self.task_tree.bind("<Button-4>", self.on_mousewheel_linux)
            self.task_tree.bind("<Button-5>", self.on_mousewheel_linux)

    def on_mousewheel_windows(self, event):
        """Handles mousewheel scrolling on Windows."""
        self.task_tree.yview_scroll(int(-1*(event.delta/120)), "units")

    def on_mousewheel_macos(self, event):
        """Handles mousewheel scrolling on macOS."""
        self.task_tree.yview_scroll(event.delta, "units")

    def on_mousewheel_linux(self, event):
        """Handles mousewheel scrolling on Linux."""
        if event.num == 4:
            self.task_tree.yview_scroll(-1, "units")
        elif event.num == 5:
            self.task_tree.yview_scroll(1, "units")

    def run(self): # ADD THIS run METHOD
        """Starts the Tkinter main event loop."""
        self.root.mainloop()

    def view_deleted_tasks(self):
        """Show a dialog with all deleted tasks for recovery or permanent deletion."""
        # Get all deleted tasks
        deleted_tasks = self.task_manager.get_filtered_tasks(show_completed=True, show_deleted=True)
        deleted_tasks = [t for t in deleted_tasks if t.get("deleted", False)]
        
        if not deleted_tasks:
            messagebox.showinfo("Deleted Tasks", "No deleted tasks found")
            return
            
        # Create a new toplevel window
        deleted_window = tk.Toplevel(self.root)
        deleted_window.title("Deleted Tasks")
        deleted_window.geometry("800x500")
        deleted_window.transient(self.root)
        deleted_window.grab_set()
        
        # Create main frame
        main_frame = ttk.Frame(deleted_window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Add info label
        ttk.Label(main_frame, 
                  text="These tasks have been deleted but can be recovered. Select tasks and use the buttons below.",
                  wraplength=780).pack(pady=(0, 10))
        
        # Create treeview for deleted tasks
        columns = ("id", "title", "due_date", "priority", "category", "assigned_to")
        tree = ttk.Treeview(main_frame, columns=columns, show='headings')
        
        # Set column headings
        tree.heading('id', text='ID')
        tree.heading('title', text='Title')
        tree.heading('due_date', text='Due Date')
        tree.heading('priority', text='Priority')
        tree.heading('category', text='Category')
        tree.heading('assigned_to', text='Assigned To')
        
        # Set column widths
        tree.column('id', width=50, anchor=tk.CENTER)
        tree.column('title', width=250)
        tree.column('due_date', width=100, anchor=tk.CENTER)
        tree.column('priority', width=80, anchor=tk.CENTER)
        tree.column('category', width=100)
        tree.column('assigned_to', width=150)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack tree and scrollbar
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Add button frame
        button_frame = ttk.Frame(deleted_window, padding=10)
        button_frame.pack(fill=tk.X)
        
        # Populate the tree
        for task in deleted_tasks:
            tree.insert('', tk.END, values=(
                task['id'],
                task['title'],
                task.get('due_date', ''),
                task.get('priority', ''),
                task.get('category', ''),
                task.get('assigned_to', '')
            ))
        
        # Function to recover selected tasks
        def recover_selected():
            selected_items = tree.selection()
            if not selected_items:
                messagebox.showinfo("Recover Task", "Please select at least one task to recover")
                return
                
            recovered_count = 0
            for item in selected_items:
                task_id = int(tree.item(item, 'values')[0])
                if self.task_manager.recover_task(task_id):
                    recovered_count += 1
                    
            if recovered_count > 0:
                self.refresh_task_list()
                self.set_status(f"Recovered {recovered_count} task(s)")
                # Refresh the deleted tasks window
                for item in tree.get_children():
                    tree.delete(item)
                    
                # Re-populate with remaining deleted tasks
                remaining_deleted = [t for t in self.task_manager.get_filtered_tasks(
                    show_completed=True, show_deleted=True) if t.get("deleted", False)]
                    
                for task in remaining_deleted:
                    tree.insert('', tk.END, values=(
                        task['id'],
                        task['title'],
                        task.get('due_date', ''),
                        task.get('priority', ''),
                        task.get('category', ''),
                        task.get('assigned_to', '')
                    ))
                    
                if not remaining_deleted:
                    deleted_window.destroy()
        
        # Function to permanently delete selected tasks
        def permanently_delete_selected():
            selected_items = tree.selection()
            if not selected_items:
                messagebox.showinfo("Delete Task", "Please select at least one task to permanently delete")
                return
                
            task_count = len(selected_items)
            confirm = messagebox.askyesno(
                "Confirm Permanent Deletion",
                f"Are you sure you want to permanently delete {task_count} task(s)?\n\n"
                "This action cannot be undone.",
                icon=messagebox.WARNING
            )
            
            if not confirm:
                return
                
            deleted_count = 0
            for item in selected_items:
                task_id = int(tree.item(item, 'values')[0])
                if self.task_manager.permanently_delete_task(task_id):
                    deleted_count += 1
                    
            if deleted_count > 0:
                self.set_status(f"Permanently deleted {deleted_count} task(s)")
                # Refresh the deleted tasks window
                for item in tree.get_children():
                    tree.delete(item)
                    
                # Re-populate with remaining deleted tasks
                remaining_deleted = [t for t in self.task_manager.get_filtered_tasks(
                    show_completed=True, show_deleted=True) if t.get("deleted", False)]
                    
                for task in remaining_deleted:
                    tree.insert('', tk.END, values=(
                        task['id'],
                        task['title'],
                        task.get('due_date', ''),
                        task.get('priority', ''),
                        task.get('category', ''),
                        task.get('assigned_to', '')
                    ))
                    
                if not remaining_deleted:
                    deleted_window.destroy()
        
        # Function to recover all tasks
        def recover_all():
            confirm = messagebox.askyesno(
                "Confirm Recovery",
                "Are you sure you want to recover all deleted tasks?",
                icon=messagebox.QUESTION
            )
            
            if not confirm:
                return
                
            recovered = self.task_manager.recover_all_deleted_tasks()
            if recovered > 0:
                self.refresh_task_list()
                self.set_status(f"Recovered all {recovered} deleted task(s)")
                deleted_window.destroy()
        
        # Function to permanently delete all tasks
        def permanently_delete_all():
            confirm = messagebox.askyesno(
                "Confirm Permanent Deletion",
                "Are you sure you want to PERMANENTLY DELETE ALL deleted tasks?\n\n"
                "This action cannot be undone.",
                icon=messagebox.WARNING
            )
            
            if not confirm:
                return
                
            deleted = self.task_manager.permanently_delete_all_deleted_tasks()
            if deleted > 0:
                self.set_status(f"Permanently deleted all {deleted} task(s)")
                deleted_window.destroy()
        
        # Add buttons
        ttk.Button(button_frame, text="Recover Selected", command=recover_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Permanently Delete Selected", command=permanently_delete_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Recover All", command=recover_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Permanently Delete All", command=permanently_delete_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Close", command=deleted_window.destroy).pack(side=tk.RIGHT, padx=5)

    def clean_expired_locks(self):
        """Clean up expired locks from the database."""
        try:
            current_time = datetime.datetime.now()
            self.cursor.execute(
                "UPDATE Tasks SET locked_by = NULL, lock_expiry = NULL WHERE lock_expiry IS NOT NULL AND lock_expiry < %s",
                (current_time,)
            )
            self.conn.commit()
            logging.info("Cleaned up expired locks")
        except Exception as e:
            logging.error(f"Error cleaning expired locks: {str(e)}", exc_info=True)

    # Call this method periodically from TaskManagerApp
    def _start_lock_cleanup_timer(self):
        """Start a timer to periodically clean up expired locks."""
        self.clean_expired_locks()
        # Schedule to run every 5 minutes (300000 ms)
        self.root.after(300000, self._start_lock_cleanup_timer)


class TaskDialog:
    """Dialog for adding or editing a task."""
    
    def __init__(self, parent, title, task=None):
        """Initialize the dialog."""
        self.result = None
        self.parent = parent
        
        # Create dialog window but keep it hidden until fully built
        self.dialog = tk.Toplevel(parent)
        self.dialog.withdraw()  # Hide the window during setup
        self.dialog.title(title)
        
        # Set fixed dialog size
        self.dialog_width = 630
        self.dialog_height = 630
        self.dialog.geometry(f"{self.dialog_width}x{self.dialog_height}")
        self.dialog.minsize(self.dialog_width, self.dialog_height)
        self.dialog.maxsize(self.dialog_width, self.dialog_height)  # Prevent resizing
        self.dialog.transient(parent)
        
        # Create main container with fixed padding
        main_container = ttk.Frame(self.dialog)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create canvas with scrollbar - use exact width calculation
        canvas_width = self.dialog_width - 40  # Account for dialog padding (20px on each side)
        self.canvas = tk.Canvas(main_container, width=canvas_width, highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=self.canvas.yview)
        
        # Create a frame inside the canvas for the form - use exact width
        form_width = canvas_width - 20  # Account for scrollbar width and some padding
        self.form_frame = ttk.Frame(self.canvas, width=form_width)
        
        # Pack scrollbar and canvas
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Configure canvas - use exact coordinates to prevent shifting
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas_frame = self.canvas.create_window((0, 0), window=self.form_frame, anchor=tk.NW, width=form_width)
        
        # Create form elements
        self.create_form(self.form_frame, task)
        
        # Create button frame at the bottom (outside the scrollable area)
        button_frame = ttk.Frame(self.dialog)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(button_frame, text="Cancel", command=self.cancel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Save", command=self.save).pack(side=tk.RIGHT, padx=5)
        
        # Bind canvas configuration and mouse wheel
        self.form_frame.bind('<Configure>', self.on_frame_configure)
        self.canvas.bind('<Configure>', self.on_canvas_configure)
        
        # Use a more reliable way to bind the mousewheel
        self.bind_mousewheel()
        
        # Bind dialog close event to ensure proper cleanup
        self.dialog.protocol("WM_DELETE_WINDOW", self.cancel)
        
        # Position the dialog
        self.position_dialog()
        
        # Pre-configure the canvas scroll region before showing the dialog
        self.dialog.update_idletasks()
        self.on_frame_configure()
        
        # Now show the dialog
        self.dialog.deiconify()
        
        # Grab focus and wait
        self.dialog.grab_set()
        self.dialog.focus_force()
        self.dialog.wait_window()
    
    def bind_mousewheel(self):
        """Bind mousewheel events to the canvas."""
        # Windows and macOS use different events and deltas
        if sys.platform.startswith('win'):
            self.canvas.bind_all("<MouseWheel>", self.on_mousewheel_windows)
        elif sys.platform.startswith('darwin'):
            self.canvas.bind_all("<MouseWheel>", self.on_mousewheel_macos)
        else:
            # Linux
            self.canvas.bind_all("<Button-4>", self.on_mousewheel_linux)
            self.canvas.bind_all("<Button-5>", self.on_mousewheel_linux)
    
    def unbind_mousewheel(self):
        """Unbind mousewheel events when dialog is closed."""
        try:
            # Unbind all mousewheel events
            self.canvas.unbind_all("<MouseWheel>")
            self.canvas.unbind_all("<Button-4>")
            self.canvas.unbind_all("<Button-5>")
        except:
            pass  # Ignore errors if the canvas is already destroyed
    
    def on_mousewheel_windows(self, event):
        """Handle mousewheel scrolling on Windows."""
        try:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except:
            pass  # Ignore errors if the canvas is destroyed
    
    def on_mousewheel_macos(self, event):
        """Handle mousewheel scrolling on macOS."""
        try:
            self.canvas.yview_scroll(int(-1 * event.delta), "units")
        except:
            pass  # Ignore errors if the canvas is destroyed
    
    def on_mousewheel_linux(self, event):
        """Handle mousewheel scrolling on Linux."""
        try:
            if event.num == 4:
                self.canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.canvas.yview_scroll(1, "units")
        except:
            pass  # Ignore errors if the canvas is destroyed
    
    def on_frame_configure(self, event=None):
        """Reset the scroll region to encompass the inner frame"""
        try:
            # Set the scroll region to the entire canvas
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            
            # Ensure the form frame stays at the left edge
            self.canvas.coords(self.canvas_frame, 0, 0)
        except:
            pass  # Ignore errors if the canvas is destroyed
    
    def on_canvas_configure(self, event):
        """When the canvas is resized, resize the inner frame to match"""
        try:
            # Maintain a fixed width for the inner frame
            width = event.width
            self.canvas.itemconfig(self.canvas_frame, width=width)
            
            # Ensure the form frame stays at the left edge
            self.canvas.coords(self.canvas_frame, 0, 0)
        except:
            pass  # Ignore errors if the canvas is destroyed
    
    def cancel(self):
        """Cancel the dialog and clean up resources."""
        # Unbind mousewheel events before destroying the dialog
        self.unbind_mousewheel()
        self.dialog.destroy()
    
    def position_dialog(self):
        """Position the dialog on the same monitor as the parent window."""
        # Update the dialog's geometry
        self.dialog.update_idletasks()
        
        # Get parent window geometry
        parent_x = self.parent.winfo_rootx()
        parent_y = self.parent.winfo_rooty()
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()
        
        # Calculate position
        x = parent_x + (parent_width - self.dialog_width) // 2
        y = parent_y + (parent_height - self.dialog_height) // 2
        
        # Set the position
        self.dialog.geometry(f"+{x}+{y}")
    
    def create_form(self, parent, task=None):
        """Create the form fields in a single frame."""
        # Configure grid
        parent.columnconfigure(0, weight=0, minsize=120)  # Label column - fixed width
        parent.columnconfigure(1, weight=1)  # Entry column - expandable
        parent.columnconfigure(2, weight=0, minsize=20)  # Button column - fixed width
        
        # Request No. and Equipment Name
        row = 0
        ttk.Label(parent, text="Request No.:").grid(row=row, column=0, sticky=tk.W, pady=5, padx=(0,5))
        self.request_no_var = tk.StringVar(value=task.get("request_no", "") if task else "")
        ttk.Entry(parent, textvariable=self.request_no_var).grid(row=row, column=1, sticky=tk.EW, pady=5)
        
        row += 1
        ttk.Label(parent, text="Equipment Name:*").grid(row=row, column=0, sticky=tk.W, pady=5, padx=(0,5))
        self.title_var = tk.StringVar(value=task["title"] if task else "")
        ttk.Entry(parent, textvariable=self.title_var).grid(row=row, column=1, sticky=tk.EW, pady=5)
        
        # Vessel Information
        row += 1
        ttk.Separator(parent, orient=tk.HORIZONTAL).grid(row=row, column=0, columnspan=2, sticky=tk.EW, pady=10)
        
        row += 1
        ttk.Label(parent, text="Applied Vessel:").grid(row=row, column=0, sticky=tk.W, pady=5, padx=(0,5))
        self.applied_vessel_var = tk.StringVar(value=task.get("applied_vessel", "") if task else "")
        ttk.Entry(parent, textvariable=self.applied_vessel_var).grid(row=row, column=1, sticky=tk.EW, pady=5)
        
        row += 1
        ttk.Label(parent, text="Rev:").grid(row=row, column=0, sticky=tk.W, pady=5, padx=(0,5))
        self.rev_var = tk.StringVar(value=task.get("rev", "") if task else "")
        ttk.Entry(parent, textvariable=self.rev_var).grid(row=row, column=1, sticky=tk.EW, pady=5)
        
        row += 1
        ttk.Label(parent, text="Drawing No.:").grid(row=row, column=0, sticky=tk.W, pady=5, padx=(0,5))
        self.drawing_no_var = tk.StringVar(value=task.get("drawing_no", "") if task else "")
        ttk.Entry(parent, textvariable=self.drawing_no_var).grid(row=row, column=1, sticky=tk.EW, pady=5)
        
        # Dates
        row += 1
        ttk.Separator(parent, orient=tk.HORIZONTAL).grid(row=row, column=0, columnspan=2, sticky=tk.EW, pady=10)
        
        row += 1
        # Create a frame for dates that spans both columns and uses grid
        date_frame = ttk.Frame(parent)
        date_frame.grid(row=row, column=0, columnspan=2, sticky=tk.EW, pady=5)
        date_frame.columnconfigure(1, weight=1)  # Make the date entries expand
        
        # Configure date frame columns for proper alignment
        date_frame.columnconfigure(0, minsize=120)  # Same as label column
        date_frame.columnconfigure(1, weight=1)     # First date picker
        date_frame.columnconfigure(2, minsize=80)   # Second label
        date_frame.columnconfigure(3, weight=1)     # Second date picker
        date_frame.columnconfigure(4, minsize=80)   # Third label
        date_frame.columnconfigure(5, weight=1)     # Third date picker
        
        # Add date pickers with proper alignment
        ttk.Label(date_frame, text="Requested Date:").grid(row=0, column=0, sticky=tk.W)
        self.requested_date_picker = DateEntry(date_frame, width=12, date_pattern='yyyy-mm-dd')
        self.requested_date_picker.grid(row=0, column=1, sticky=tk.EW, padx=(0,10))
        
        ttk.Label(date_frame, text="Started date:").grid(row=0, column=2, sticky=tk.W)
        self.date_started_picker = DateEntry(date_frame, width=12, date_pattern='yyyy-mm-dd')
        self.date_started_picker.grid(row=0, column=3, sticky=tk.EW, padx=(0,10))
        
        ttk.Label(date_frame, text="Due Date:").grid(row=0, column=4, sticky=tk.W)
        self.due_date_picker = DateEntry(date_frame, width=12, date_pattern='yyyy-mm-dd')
        self.due_date_picker.grid(row=0, column=5, sticky=tk.EW)
        
        # Set existing dates if available
        if task:
            for picker, date_key in [(self.requested_date_picker, "requested_date"),
                                   (self.date_started_picker, "date_started"),
                                   (self.due_date_picker, "due_date")]:
                if task.get(date_key):
                    try:
                        date_value = datetime.datetime.strptime(task[date_key], "%Y-%m-%d").date()
                        picker.set_date(date_value)
                    except ValueError:
                        pass
        
        # Links
        row += 1
        ttk.Separator(parent, orient=tk.HORIZONTAL).grid(row=row, column=0, columnspan=2, sticky=tk.EW, pady=10)
        
        row += 1
        ttk.Label(parent, text="Link:").grid(row=row, column=0, sticky=tk.W, pady=5)
        
        # Create a frame for the link field and browse button
        link_frame = ttk.Frame(parent)
        link_frame.grid(row=row, column=1, sticky=tk.EW, pady=5)
        link_frame.columnconfigure(0, weight=1)  # Make the entry expand
        
        self.link_var = tk.StringVar(value=task.get("link", "") if task else "")
        
        # Create a custom style for the browse button to match entry height
        # FIX: Use the style from the parent window instead of the dialog
        style = ttk.Style()
        style.configure("BrowseButton.TButton", padding=0)
        
        # Create the entry
        link_entry = ttk.Entry(link_frame, textvariable=self.link_var)
        link_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        # Add browse button - using pack instead of grid for better height matching
        browse_btn = ttk.Button(link_frame, text="Browse", command=self.browse_folder, style="BrowseButton.TButton")
        browse_btn.pack(side=tk.RIGHT, fill=tk.Y)
        
        row += 1
        ttk.Label(parent, text="SDB Link:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.sdb_link_var = tk.StringVar(value=task.get("sdb_link", "") if task else "")
        sdb_link_entry = ttk.Entry(parent, textvariable=self.sdb_link_var)
        sdb_link_entry.grid(row=row, column=1, sticky=tk.EW, pady=5)
        
        # Description
        row += 1
        ttk.Separator(parent, orient=tk.HORIZONTAL).grid(row=row, column=0, columnspan=2, sticky=tk.EW, pady=10)
        
        row += 1
        ttk.Label(parent, text="Description:").grid(row=row, column=0, sticky=tk.W, pady=5, padx=(0,5))
        
        # Create a frame to contain both text widget and scrollbar in the same row as the label
        desc_frame = ttk.Frame(parent)
        desc_frame.grid(row=row, column=1, sticky=tk.EW, pady=5)
        desc_frame.columnconfigure(0, weight=1)  # Make text widget expand
        
        self.description_text = tk.Text(desc_frame, height=4, wrap=tk.WORD)
        self.description_text.grid(row=0, column=0, sticky=tk.EW)
        if task and task.get("description"):
            self.description_text.insert("1.0", task["description"])
        
        # Add scrollbar to description inside the frame
        desc_scrollbar = ttk.Scrollbar(desc_frame, orient="vertical", command=self.description_text.yview)
        desc_scrollbar.grid(row=0, column=1, sticky=tk.NS)
        self.description_text.configure(yscrollcommand=desc_scrollbar.set)
        
        # Priority
        row += 1
        ttk.Label(parent, text="Priority:").grid(row=row, column=0, sticky=tk.W, pady=5)
        priority_frame = ttk.Frame(parent)
        priority_frame.grid(row=row, column=1, sticky=tk.W, pady=5)
        
        self.priority_var = tk.StringVar(value=task["priority"] if task else "medium")
        ttk.Radiobutton(priority_frame, text="Low", variable=self.priority_var, value="low").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(priority_frame, text="Medium", variable=self.priority_var, value="medium").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(priority_frame, text="High", variable=self.priority_var, value="high").pack(side=tk.LEFT, padx=5)
        
        # Category
        row += 1
        ttk.Label(parent, text="Category:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.category_var = tk.StringVar(value=task["category"] if task else "general")
        ttk.Entry(parent, textvariable=self.category_var).grid(row=row, column=1, sticky=tk.EW, pady=5)
        
        # Main Staff
        row += 1
        ttk.Label(parent, text="Main Staff:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.main_staff_var = tk.StringVar(value=task.get("main_staff", "") if task else "")
        self.main_staff_combo = ttk.Combobox(parent, textvariable=self.main_staff_var, 
                                           values=USERS[1:], state="readonly")
        self.main_staff_combo.grid(row=row, column=1, sticky=tk.EW, pady=5)
        self.disable_combobox_mousewheel(self.main_staff_combo)  # Disable mousewheel scrolling
        row += 1
        
        # Assigned To
        ttk.Label(parent, text="Assigned To:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.assigned_to_var = tk.StringVar(value=task.get("assigned_to", "") if task else "")
        self.assigned_to_combo = ttk.Combobox(parent, textvariable=self.assigned_to_var, 
                                            values=USERS[1:], state="readonly")
        self.assigned_to_combo.grid(row=row, column=1, sticky=tk.EW, pady=5)
        self.disable_combobox_mousewheel(self.assigned_to_combo)  # Disable mousewheel scrolling
        row += 1
        
        # Qtd Mhr
        ttk.Label(parent, text="Qtd Mhr:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.qtd_mhr_var = tk.StringVar(value=str(task.get("qtd_mhr", "")) if task and task.get("qtd_mhr") is not None else "") # Ensure default is empty string
        ttk.Entry(parent, textvariable=self.qtd_mhr_var).grid(row=row, column=1, sticky=tk.EW, pady=5)
        row += 1

        # Actual Mhr
        ttk.Label(parent, text="Actual Mhr:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.actual_mhr_var = tk.StringVar(value=str(task.get("actual_mhr", "")) if task and task.get("actual_mhr") is not None else "") # Ensure default is empty string
        ttk.Entry(parent, textvariable=self.actual_mhr_var).grid(row=row, column=1, sticky=tk.EW, pady=5)
        row += 1

        # Status (only for editing)
        if task is not None:
            row += 1
            ttk.Label(parent, text="Status:").grid(row=row, column=0, sticky=tk.W, pady=5)
            self.completed_var = tk.BooleanVar(value=task["completed"])
            ttk.Checkbutton(parent, text="Completed", variable=self.completed_var).grid(row=row, column=1, sticky=tk.W, pady=5)
    
    def disable_combobox_mousewheel(self, combobox):
        """Disable mousewheel scrolling for a combobox to prevent accidental changes."""
        def block_mousewheel(event):
            return "break"
        
        # Bind mousewheel events for different platforms
        combobox.bind("<MouseWheel>", block_mousewheel)  # Windows
        combobox.bind("<Button-4>", block_mousewheel)    # Linux
        combobox.bind("<Button-5>", block_mousewheel)    # Linux
        # Remove the incorrect macOS binding that was causing the error
    
    def browse_folder(self):
        """Browse for a folder and set it as the link."""
        # Get the applied vessel value
        applied_vessel = self.applied_vessel_var.get().strip()
        
        # Check if applied vessel is provided
        if not applied_vessel:
            messagebox.showwarning("Missing Information", 
                                 "Please enter a value for Applied Vessel before browsing.",
                                 parent=self.dialog)
            return
        
        # Define the base project path
        base_path = r"\\srb096154\01_CESSD_SCG_CAD\01_Projects"
        vessel_path = os.path.join(base_path, applied_vessel)
        
        # Check if the vessel folder exists
        if not os.path.exists(vessel_path):
            # Ask if the user wants to create the folder
            create_folder = messagebox.askyesno(
                "Create Folder",
                f"Folder for vessel '{applied_vessel}' does not exist.\n\nDo you want to create it?",
                parent=self.dialog
            )
            
            if create_folder:
                try:
                    # Create the folder
                    os.makedirs(vessel_path, exist_ok=True)
                except Exception as e:
                    messagebox.showerror(
                        "Error",
                        f"Could not create folder: {vessel_path}\n\nError: {str(e)}",
                        parent=self.dialog
                    )
                    return
            else:
                return
        
        # Open folder dialog
        try:
            from tkinter import filedialog
            
            # Use the vessel path as the initial directory if it exists
            initial_dir = vessel_path if os.path.exists(vessel_path) else os.path.dirname(vessel_path)
            
            folder_path = filedialog.askdirectory(
                title="Select Folder",
                initialdir=initial_dir,
                parent=self.dialog
            )
            
            # Update the link field if a folder was selected
            if folder_path:
                self.link_var.set(folder_path)
        except Exception as e:
            messagebox.showerror(
                "Error",
                f"An error occurred while browsing for folders: {str(e)}",
                parent=self.dialog
            )
    
    def test_link(self, link):
        """Test the link by opening it in a web browser."""
        if not link:
            messagebox.showinfo("Information", "Please enter a link first.")
            return
        
        # Use the open_link method from the parent application
        try:
            import webbrowser
            
            # Add http:// prefix if not present and not a local file path
            if not link.startswith(("http://", "https://", "file://", "www.", "/")):
                # Check if it might be a local file path
                if os.path.exists(link) or link.startswith(("C:", "D:", "E:", "F:", "\\")):
                    # It's likely a local file path
                    link = f"file:///{link.replace('\\', '/')}"
                else:
                    # Assume it's a web URL
                    link = f"http://{link}"
            
            webbrowser.open(link)
            messagebox.showinfo("Success", f"Link opened successfully: {link}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open link: {link}\n\nError: {str(e)}")
    
    def save(self):
        """Save the form data."""
        # Validate required fields
        validation_errors = []
        
        # Request No validation
        request_no = self.request_no_var.get().strip()
        if not request_no:
            validation_errors.append("Request No. is required")
        
        # Equipment Name validation
        title = self.title_var.get().strip()
        if not title:
            validation_errors.append("Equipment Name is required")
        
        # Applied Vessel validation
        applied_vessel = self.applied_vessel_var.get().strip()
        if not applied_vessel:
            validation_errors.append("Applied Vessel is required")
        
        # Rev validation
        rev = self.rev_var.get().strip()
        if not rev:
            validation_errors.append("Rev is required")
        else:
            if not rev.isdigit(): # Check if rev is numeric
                validation_errors.append("Rev must be a valid integer")
            else:
                try:
                    rev_num = int(rev)
                    if rev_num < 0: # Check if rev is non-negative
                        validation_errors.append("Rev must be a non-negative integer")
                    else:
                        self.rev_var.set(str(rev_num)) # Store the converted integer back
                except ValueError: # Redundant, but kept for clarity
                    validation_errors.append("Rev must be a valid integer")
        
        # Link validation
        link = self.link_var.get().strip()
        if not link:
            validation_errors.append("Link is required")
        
        # SDB Link validation
        sdb_link = self.sdb_link_var.get().strip()
        if not sdb_link:
            validation_errors.append("SDB Link is required")
        
        # Main Staff validation
        main_staff = self.main_staff_var.get().strip()
        if not main_staff:
            validation_errors.append("Main Staff is required")
        
        # Priority validation
        priority = self.priority_var.get().lower()
        if priority not in ["low", "medium", "high"]:
            validation_errors.append("Priority must be one of: Low, Medium, High")
        
        # Category validation (example: alphanumeric and spaces only)
        category = self.category_var.get().strip()
        if not category:
            validation_errors.append("Category is required")
        elif not category.replace(" ", "").isalnum(): # Allow alphanumeric and spaces
            validation_errors.append("Category must be alphanumeric and spaces only")
        
        # Main Staff and Assigned To validation (check against USERS list)
        main_staff = self.main_staff_var.get()
        if main_staff and main_staff not in USERS[1:]: # USERS[1:] to exclude "All"
            validation_errors.append(f"Main Staff must be one of: {', '.join(USERS[1:])}")

        assigned_to = self.assigned_to_var.get()
        if assigned_to and assigned_to not in USERS[1:]: # USERS[1:] to exclude "All"
            validation_errors.append(f"Assigned To must be one of: {', '.join(USERS[1:])}")
        
        # Qtd Mhr validation
        qtd_mhr_str = self.qtd_mhr_var.get().strip()
        if qtd_mhr_str: # Only validate if not empty
            if not qtd_mhr_str.isdigit():
                validation_errors.append("Qtd Mhr must be an integer")
            else:
                try:
                    qtd_mhr = int(qtd_mhr_str)
                    if qtd_mhr < 0:
                        validation_errors.append("Qtd Mhr must be a non-negative integer")
                except ValueError: # Redundant, but kept for clarity
                    validation_errors.append("Qtd Mhr must be an integer")
        else:
            qtd_mhr = 0 # Default to 0 if empty

        # Actual Mhr validation
        actual_mhr_str = self.actual_mhr_var.get().strip()
        if actual_mhr_str: # Only validate if not empty
            if not actual_mhr_str.isdigit():
                validation_errors.append("Actual Mhr must be an integer")
            else:
                try:
                    actual_mhr = int(actual_mhr_str)
                    if actual_mhr < 0:
                        validation_errors.append("Actual Mhr must be a non-negative integer")
                except ValueError: # Redundant, but kept for clarity
                    validation_errors.append("Actual Mhr must be an integer")
        else:
            actual_mhr = 0 # Default to 0 if empty

        # Show all validation errors if any
        if validation_errors:
            error_message = "Please correct the following errors:\n\n" + "\n".join(f"• {error}" for error in validation_errors)
            messagebox.showerror("Validation Error", error_message, parent=self.dialog)
            return
        
        # Get dates - ensure we get the raw string value from the picker
        try:
            requested_date = self.requested_date_picker.get()
            if requested_date:
                # Store as string in YYYY-MM-DD format
                requested_date = requested_date
        except Exception as e:
            print(f"Error getting requested date: {str(e)}")
            requested_date = None
            
        try:
            date_started = self.date_started_picker.get()
            if date_started:
                # Store as string in YYYY-MM-DD format
                date_started = date_started
        except Exception as e:
            print(f"Error getting start date: {str(e)}")
            date_started = None
            
        try:
            due_date = self.due_date_picker.get()
            if due_date:
                # Store as string in YYYY-MM-DD format
                due_date = due_date
        except Exception as e:
            print(f"Error getting due date: {str(e)}")
            due_date = None
        
        # Create result with all fields
        self.result = {
            "request_no": request_no,
            "title": title,
            "description": self.description_text.get("1.0", tk.END).strip() or None,  # Make None if empty
            "requested_date": requested_date,
            "date_started": date_started,
            "due_date": due_date,
            "priority": self.priority_var.get(),
            "category": self.category_var.get().strip() or "general",
            "main_staff": main_staff,
            "assigned_to": self.assigned_to_var.get().strip() or None,
            "applied_vessel": applied_vessel,
            "rev": rev,
            "drawing_no": self.drawing_no_var.get().strip() or None,  # Make None if empty
            "link": link,
            "sdb_link": sdb_link,
            "qtd_mhr": qtd_mhr if qtd_mhr_str else 0, # Use validated integer value, default to 0 if empty
            "actual_mhr": actual_mhr if actual_mhr_str else 0, # Use validated integer value, default to 0 if empty
            "request_no": request_no,
            "requested_date": requested_date,
            "date_started": date_started
        }
        
        # Add completed status if editing
        if hasattr(self, "completed_var"):
            self.result["completed"] = self.completed_var.get()
        
        # Unbind mousewheel events before destroying the dialog
        self.unbind_mousewheel()
        
        # Close dialog
        self.dialog.destroy()


class TaskManager:
    """Class to manage daily tasks."""
    
    def __init__(self):
        """Initialize the task manager."""
        # Database connection info
        self.server = "10.195.102.56"  # Server address (IP address)
        self.database = "TaskManagerDB"
        self.user = None  # Will be set by _get_user_credentials
        self.password = None  # Will be set by _get_user_credentials
        
        # Initialize internal state
        self._tasks = []  # List of tasks loaded from the database
        self._task_id_counter = 1
        self._deleted_tasks = []  # Track recently deleted tasks for undo
        self._transaction_active = False
        
        # Create standardized connection references
        self.conn = None  # Main connection object
        self.cursor = None  # Main cursor object
        
        # Initialize everything
        self._get_user_credentials()  # Set up credentials
        self._connect()  # Connect to database
        self._ensure_deleted_field_exists()
        self._ensure_lock_fields_exist()  # Add this line
        
        # Load tasks after schema updates
        self._load_tasks()
    
    def _ensure_deleted_field_exists(self):
        """Ensure the 'deleted' field exists in the Tasks table."""
        try:
            # Check if the deleted column exists - using SQL Server syntax
            self.cursor.execute("""
                SELECT COUNT(*) 
                FROM INFORMATION_SCHEMA.COLUMNS 
                WHERE TABLE_NAME = 'Tasks' 
                AND COLUMN_NAME = 'deleted'
            """)
            
            column_exists = self.cursor.fetchone()[0] > 0
            
            if not column_exists:
                # Add the column if it doesn't exist
                self.cursor.execute("""
                    ALTER TABLE Tasks 
                    ADD deleted BIT NOT NULL DEFAULT 0
                """)
                self.conn.commit()
                logging.info("Added 'deleted' column to Tasks table")
            
            # Also check if version column exists - using SQL Server syntax
            self.cursor.execute("""
                SELECT COUNT(*) 
                FROM INFORMATION_SCHEMA.COLUMNS 
                WHERE TABLE_NAME = 'Tasks' 
                AND COLUMN_NAME = 'version'
            """)
            
            version_exists = self.cursor.fetchone()[0] > 0
            
            if not version_exists:
                # Add version column for optimistic concurrency control
                self.cursor.execute("""
                    ALTER TABLE Tasks 
                    ADD version INT NOT NULL DEFAULT 1
                """)
                self.conn.commit()
                logging.info("Added 'version' column to Tasks table")
            
        except pymssql.Error as e:
            logging.error(f"Error checking/adding columns: {str(e)}", exc_info=True)
            # We'll continue even if this fails
    
    def _get_user_credentials(self):
        """Retrieves SQL Server credentials based on the current user."""
        username = getpass.getuser()  # Get the current logged-in username

        # Set credentials based on username
        if username == "a0011071":
            self.user = "TaskUser1"
            self.password = "pass1"
        elif username == "a0010756":
            self.user = "TaskUser2"
            self.password = "pass1"
        elif username == "a0012923":
            self.user = "TaskUser3"
            self.password = "pass1"
        elif username == "a0010751":
            self.user = "TaskUser4"
            self.password = "pass1"
        elif username == "a0012501":
            self.user = "TaskUser5"
            self.password = "pass1"
        elif username == "a0008432":
            self.user = "TaskUser6"
            self.password = "pass1"
        elif username == "a0003878":
            self.user = "TaskUser7"
            self.password = "pass1"
        else:
            # Use default credentials for other users
            print(f"Warning: Username '{username}' is not recognized as a specific user. Using default login.")
            self.user = "DefaultFallbackUser"
            self.password = "pass1"

        # Store sql_username as instance attribute - not needed with this approach
        # self.sql_username = sql_username
    
    def _connect(self):
        """Connect to the SQL Server database."""
        try:
            # Close existing connection if any
            if self.conn:
                try:
                    self.conn.close()
                except:
                    pass
            
            # Connect to the database
            self.conn = pymssql.connect(
                server=self.server,
                database=self.database,
                user=self.user,
                password=self.password
            )
            
            # Create a cursor
            self.cursor = self.conn.cursor()
            
            logging.info("Successfully connected to the SQL Server database.")
        except pymssql.Error as e:
            logging.error(f"Error connecting to database: {str(e)}", exc_info=True)
            messagebox.showerror("Database Connection Error", 
                                f"Could not connect to the database.\n\nError: {str(e)}")
            raise
    
    def _load_tasks(self) -> List[Dict[str, Any]]:
        """Load tasks from the SQL Server database."""
        try:
            self.cursor.execute("SELECT * FROM [dbo].[Tasks]")
            rows = self.cursor.fetchall()
            
            # Convert rows to a list of dictionaries
            tasks = []
            for row in rows:
                task = {
                    "id": row[0],
                    "title": row[1],
                    "description": row[2],
                    "created_date": row[3].strftime("%Y-%m-%d") if isinstance(row[3], datetime.datetime) else None,
                    "created_by": row[4],
                    "due_date": row[5] if isinstance(row[5], str) else
                               row[5].strftime("%Y-%m-%d") if row[5] else None,
                    "priority": row[6],
                    "category": row[7],
                    "main_staff": row[8],
                    "assigned_to": row[9],
                    "completed": row[10],
                    "applied_vessel": row[11],
                    "rev": row[12],
                    "drawing_no": row[13],
                    "link": row[14],
                    "sdb_link": row[15],
                    "request_no": row[16],
                    "requested_date": row[17] if isinstance(row[17], str) else
                                    row[17].strftime("%Y-%m-%d") if row[17] else None,
                    "date_started": row[18] if isinstance(row[18], str) else
                                  row[18].strftime("%Y-%m-%d") if row[18] else None,
                    "qtd_mhr": row[19],
                    "actual_mhr": row[20],
                    "last_modified": row[21].strftime("%Y-%m-%d %H:%M:%S") if isinstance(row[21], datetime.datetime) else None,
                    "modified_by": row[22],
                    # Get the deleted field if it exists, otherwise default to False
                    "deleted": row[23] if len(row) > 23 else False
                }
                tasks.append(task)
            
            return tasks
        except pymssql.Error as e:
            logging.error("Error loading tasks from database", exc_info=True)
            messagebox.showerror("Database Error", f"Error loading tasks: {str(e)}")
            return []
    
    def _save_tasks(self) -> None:
        """Commit changes to the database."""
        # If we're not explicitly in a transaction, commit changes to the database
        if self.conn and not self._transaction_active:
            try:
                self.conn.commit()
            except pymssql.Error as e:
                logging.error("Error committing changes to database", exc_info=True)
                messagebox.showerror("Database Error", f"Error saving changes: {str(e)}")
    
    def _get_current_user(self):
        """Get the current user's username."""
        try:
            return os.getenv('USERNAME') or os.getenv('USER') or 'Unknown User'
        except:
            return 'Unknown User'
    
    def add_task(self, title: str, description: str = "", due_date: Optional[str] = None, 
                 priority: str = "medium", category: str = "general", 
                 main_staff: Optional[str] = None, assigned_to: Optional[str] = None,
                 applied_vessel: str = "", rev: str = "", drawing_no: str = "",
                 link: str = "", sdb_link: str = "", request_no: str = "",
                 requested_date: Optional[str] = None, date_started: Optional[str] = None,
                 qtd_mhr: int = 0, actual_mhr: int = 0) -> Optional[int]:
        """Add a new task with proper transaction handling."""
        self.begin_transaction()
        
        try:
            current_time = datetime.datetime.now()
            current_user = self._get_current_user()
            
            # Validate dates
            def validate_date(date_str):
                if not date_str:
                    return None
                try:
                    if isinstance(date_str, str):
                        # Try different date formats
                        for fmt in ["%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"]:
                            try:
                                return datetime.datetime.strptime(date_str, fmt).strftime("%Y-%m-%d")
                            except ValueError:
                                continue
                        # If none of the formats match
                        return None
                    else:
                        return date_str
                except Exception:
                    return None
            
            # Process date fields
            due_date = validate_date(due_date)
            requested_date = validate_date(requested_date)
            date_started = validate_date(date_started)
            
            # Prepare query with version field initialized to 1
            sql = """
            INSERT INTO Tasks (
                title, description, created_date, created_by, due_date, 
                priority, category, main_staff, assigned_to, completed,
                applied_vessel, rev, drawing_no, link, sdb_link, 
                request_no, requested_date, date_started, qtd_mhr, actual_mhr,
                last_modified, modified_by, deleted, version
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """
            
            values = (
                title, description, current_time, current_user, due_date,
                priority, category, main_staff, assigned_to, False,
                applied_vessel, rev, drawing_no, link, sdb_link,
                request_no, requested_date, date_started, qtd_mhr, actual_mhr,
                current_time, current_user, False, 1  # Initial version is 1
            )
            
            self.cursor.execute(sql, values)
            task_id = self.cursor.lastrowid
            
            self.commit_transaction()
            return task_id
        
        except Exception as e:
            print(f"Error adding task: {e}")
            self.rollback_transaction()
            return None
    
    def add_task_direct_id(self, task_id: int, title: str, description: str = "", due_date: Optional[str] = None,
                 priority: str = "medium", category: str = "general",
                 main_staff: Optional[str] = None, assigned_to: Optional[str] = None,
                 applied_vessel: str = "", rev: str = "", drawing_no: str = "",
                 link: str = "", sdb_link: str = "", request_no: str = "",
                 requested_date: Optional[str] = None, date_started: Optional[str] = None,
                 qtd_mhr: int = 0, actual_mhr: int = 0, completed: bool = False) -> None: # Added completed status
        """Add a new task to the SQL Server database with a specified ID (for undo)."""
        try:
            # Get current date and time
            current_datetime = datetime.datetime.now()
            created_date = current_datetime.strftime("%Y-%m-%d")
            created_by = self._get_current_user() # Get current user for created_by
            last_modified = created_date # Initialize last_modified for new tasks in undo
            modified_by = created_by # Initialize modified_by for new tasks in undo

            # Convert dates to strings if they are date objects
            if isinstance(due_date, datetime.date):
                due_date = due_date.strftime("%Y-%m-%d")
            if isinstance(requested_date, datetime.date):
                requested_date = requested_date.strftime("%Y-%m-%d")
            if isinstance(date_started, datetime.date):
                date_started = date_started.strftime("%Y-%m-%d")

            # SQL query to insert a new task - EXACTLY 23 columns
            sql = """
            INSERT INTO Tasks (
                id, title, description, created_date, created_by,
                due_date, priority, category, main_staff, assigned_to,
                completed, applied_vessel, rev, drawing_no, link,
                sdb_link, request_no, requested_date, date_started,
                qtd_mhr, actual_mhr, last_modified, modified_by
            )
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s) -- EXACTLY 23 placeholders
            """

            # Task data - EXACTLY 23 values in correct order
            task_data = (
                task_id, title, description, created_date, created_by,
                due_date, priority.lower(), category, main_staff, assigned_to,
                completed, applied_vessel, rev, drawing_no, link,
                sdb_link, request_no, requested_date, date_started,
                qtd_mhr, actual_mhr, last_modified, modified_by
            )

            # Double check: Ensure len(task_data) == number of placeholders == 23
            if len(task_data) != 23:
                logging.error(f"Value count mismatch in add_task_direct_id: Expected 23, got {len(task_data)}")
                raise ValueError(f"Value count mismatch in add_task_direct_id: Expected 23, got {len(task_data)}")


            # Execute the query
            self.cursor.execute(sql, task_data)
            self.conn.commit()

            # Refresh the task list
            self._tasks = self._load_tasks()

        except pymssql.Error as e:
            logging.error("Error adding task to database with direct ID", exc_info=True)
            messagebox.showerror("Database Error", f"Error adding task (undo): {str(e)}")
        except ValueError as ve:
            messagebox.showerror("Programming Error", f"Value mismatch in code: {str(ve)}")
    
    def get_filtered_tasks(self, show_completed: bool = False, category: Optional[str] = None, 
                          main_staff: Optional[str] = None, assigned_to: Optional[str] = None,
                          show_deleted: bool = False) -> List[Dict[str, Any]]:
        """Get tasks filtered by various criteria."""
        filtered_tasks = []
        
        for task in self._tasks:
            # Skip deleted tasks unless explicitly requested
            if task.get("deleted", False) and not show_deleted:
                continue
                
            # Apply other filters as before
            if not show_completed and task.get("completed", False):
                continue
                
            if category and category != "All" and task.get("category", "").lower() != category.lower():
                continue
                
            if main_staff and main_staff != "All" and task.get("main_staff", "").lower() != main_staff.lower():
                continue
                
            if assigned_to and assigned_to != "All" and task.get("assigned_to", "").lower() != assigned_to.lower():
                continue
                
            filtered_tasks.append(task)
        
        return filtered_tasks
    
    def update_task(self, task_id: int, **kwargs) -> bool:
        """
        Update a task with optimistic concurrency control.
        Returns True if update was successful, False if there was a concurrency conflict.
        """
        # Start a transaction
        self.begin_transaction()
        
        try:
            # Get current version of the task
            current_version_query = "SELECT version FROM Tasks WHERE id = %s"
            self.cursor.execute(current_version_query, (task_id,))
            result = self.cursor.fetchone()
            
            if not result:
                self.rollback_transaction()
                return False  # Task not found
                
            current_version = result[0] or 1
            
            # If client provided a version, check it matches
            if 'expected_version' in kwargs:
                expected_version = kwargs.pop('expected_version')
                if expected_version != current_version:
                    self.rollback_transaction()
                    return False  # Version mismatch - someone else modified the task
            
            # Validate date fields
            def validate_date(date_str):
                if not date_str:
                    return None
                try:
                    if isinstance(date_str, str):
                        # Try different date formats
                        for fmt in ["%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"]:
                            try:
                                return datetime.datetime.strptime(date_str, fmt).strftime("%Y-%m-%d")
                            except ValueError:
                                continue
                        # If none of the formats match
                        return None
                    else:
                        return date_str
                except Exception:
                    return None
            
            # Process date fields
            if 'due_date' in kwargs:
                kwargs['due_date'] = validate_date(kwargs['due_date'])
            if 'requested_date' in kwargs:
                kwargs['requested_date'] = validate_date(kwargs['requested_date'])
            if 'date_started' in kwargs:
                kwargs['date_started'] = validate_date(kwargs['date_started'])
                
            # Add last modified information
            kwargs['last_modified'] = datetime.datetime.now()
            kwargs['modified_by'] = self._get_current_user()
            
            # Increment version for optimistic concurrency control
            new_version = current_version + 1
            kwargs['version'] = new_version
            
            # Build SQL update statement
            sql_parts = []
            values = []
            
            for key, value in kwargs.items():
                if key in ['expected_version']:  # Skip non-database fields
                    continue
                sql_parts.append(f"{key} = %s")
                values.append(value)
            
            # Add WHERE clause with version check for optimistic concurrency control
            sql = f"UPDATE Tasks SET {', '.join(sql_parts)} WHERE id = %s AND version = %s"
            values.extend([task_id, current_version])
            
            # Execute the update
            self.cursor.execute(sql, tuple(values))
            
            # Check if update was successful
            if self.cursor.rowcount == 0:
                # No rows updated - another user must have changed the record
                self.rollback_transaction()
                return False
                
            # Commit the transaction
            self.commit_transaction()
            return True
            
        except Exception as e:
            print(f"Error updating task: {e}")
            self.rollback_transaction()
            return False
    
    def _clean_task_files(self):
        """Clean up database by removing orphaned tasks or data."""
        # This method would need to be reimplemented for SQL Server
        # For now, it's just a placeholder as this type of cleanup might be handled differently
        pass
    
    def complete_task(self, task_id: int) -> None:
        """
        Mark a task as completed.
        
        Args:
            task_id: ID of the task to mark as completed
        """
        self.update_task(task_id, completed=True)
    
    def delete_task(self, task_id: int) -> bool:
        """Soft delete a task by marking it as deleted."""
        task = self._find_task_by_id(task_id)
        if not task:
            return False
        
        # Store a copy of the task before modification for potential undo
        task_copy = task.copy()
        self._deleted_tasks.append(task_copy)
        
        # Update the task in the database
        try:
            self.cursor.execute("UPDATE [dbo].[Tasks] SET deleted = 1 WHERE id = %s", (task_id,))
            self.conn.commit()
            
            # Update local copy
            task["deleted"] = True
            
            return True
        except pymssql.Error as e:
            logging.error(f"Error soft-deleting task {task_id}", exc_info=True)
            messagebox.showerror("Database Error", f"Error deleting task: {str(e)}")
            return False

    def batch_delete_tasks(self, task_ids: List[int]) -> Tuple[List[int], List[int]]:
        """Soft delete multiple tasks by marking them as deleted."""
        if not task_ids:
            return [], []
        
        success_ids = []
        failed_ids = []
        
        try:
            # Begin a transaction for batch operations
            self.cursor.execute("BEGIN TRANSACTION")
            
            for task_id in task_ids:
                task = self._find_task_by_id(task_id)
                if task:
                    # Store a copy of the task before modification for potential undo
                    task_copy = task.copy()
                    self._deleted_tasks.append(task_copy)
                    
                    # Update the database
                    self.cursor.execute("UPDATE [dbo].[Tasks] SET deleted = 1 WHERE id = %s", (task_id,))
                    
                    # Update local copy
                    task["deleted"] = True
                    success_ids.append(task_id)
                else:
                    failed_ids.append(task_id)
            
            # Commit the transaction
            self.cursor.execute("COMMIT TRANSACTION")
            self.conn.commit()
        except pymssql.Error as e:
            # Roll back the transaction on error
            try:
                self.cursor.execute("ROLLBACK TRANSACTION")
            except:
                pass
            
            logging.error(f"Error in batch delete: {str(e)}", exc_info=True)
            messagebox.showerror("Database Error", f"Error deleting tasks: {str(e)}")
            
            # Consider all tasks as failed if there was an error
            failed_ids = task_ids
            success_ids = []
        
        return success_ids, failed_ids
    
    def batch_save(self):
        """
        Save all tasks after batch operations like multiple deletions.
        This is more efficient than saving after each deletion.
        """
        # Save all tasks to their appropriate files
        # self._save_tasks()
        pass
    
    def _find_task_by_id(self, task_id: int) -> Optional[Dict[str, Any]]:
        """Find a task by ID in the SQL Server database."""
        try:
            # SQL query to find a task by ID
            sql = "SELECT * FROM Tasks WHERE id = %s"
            
            # Execute the query
            self.cursor.execute(sql, (task_id,))
            row = self.cursor.fetchone()
            
            if row:
                # Convert row to a dictionary
                task = {
                    "id": row[0],
                    "title": row[1],
                    "description": row[2],
                    "created_date": row[3].strftime("%Y-%m-%d") if isinstance(row[3], datetime.datetime) else None,
                    "created_by": row[4],
                    "due_date": row[5] if isinstance(row[5], str) else
                               row[5].strftime("%Y-%m-%d") if row[5] else None,
                    "priority": row[6],
                    "category": row[7],
                    "main_staff": row[8],
                    "assigned_to": row[9],
                    "completed": row[10],
                    "applied_vessel": row[11],
                    "rev": row[12],
                    "drawing_no": row[13],
                    "link": row[14],
                    "sdb_link": row[15],
                    "request_no": row[16],
                    "requested_date": row[17] if isinstance(row[17], str) else
                                    row[17].strftime("%Y-%m-%d") if row[17] else None,
                    "date_started": row[18] if isinstance(row[18], str) else
                                  row[18].strftime("%Y-%m-%d") if row[18] else None,
                    "qtd_mhr": row[19],
                    "actual_mhr": row[20],
                    "last_modified": row[21].strftime("%Y-%m-%d %H:%M:%S") if isinstance(row[21], datetime.datetime) else None,
                    "modified_by": row[22],
                    # Get the deleted field if it exists, otherwise default to False
                    "deleted": row[23] if len(row) > 23 else False
                }
                return task
            else:
                return None
        except pymssql.Error as e:
            logging.error("Error finding task by ID in database", exc_info=True)
            messagebox.showerror("Database Error", f"Error finding task by ID: {str(e)}")
            return None
    
    def reload_tasks(self):
        """Reload tasks from the database."""
        # Close existing connection if any
        if self.conn:
            try:
                self.conn.close()
                self.conn = None
                self.cursor = None
            except:
                pass
            
        # Reconnect and reload
        self._connect()
        self._tasks = self._load_tasks()
        return self._tasks
    
    def begin_transaction(self):
        """Begin a database transaction with proper isolation level for SQL Server."""
        try:
            if hasattr(self, 'conn') and self.conn:
                # For SQL Server
                self.cursor.execute("SET TRANSACTION ISOLATION LEVEL SERIALIZABLE")
                self.cursor.execute("BEGIN TRANSACTION")
                self._transaction_active = True
        except Exception as e:
            logging.error(f"Error beginning transaction: {str(e)}", exc_info=True)

    def commit_transaction(self):
        """Commit the current transaction."""
        try:
            if hasattr(self, 'conn') and self.conn and self._transaction_active:
                self.conn.commit()
                self._transaction_active = False
        except Exception as e:
            logging.error(f"Error committing transaction: {str(e)}", exc_info=True)
            self.rollback_transaction()

    def rollback_transaction(self):
        """Rollback the current transaction."""
        try:
            if hasattr(self, 'conn') and self.conn and self._transaction_active:
                self.conn.rollback()
                self._transaction_active = False
        except Exception as e:
            logging.error(f"Error rolling back transaction: {str(e)}", exc_info=True)

    def get_sql_username(self):
        """Get the SQL username used for the database connection."""
        return self.user

    def recover_task(self, task_id: int) -> bool:
        """Recover a soft-deleted task."""
        task = self._find_task_by_id(task_id)
        if not task or not task.get("deleted", False):
            return False
        
        try:
            # Update the database
            self.cursor.execute("UPDATE [dbo].[Tasks] SET deleted = 0 WHERE id = %s", (task_id,))
            self.conn.commit()
            
            # Update local copy
            task["deleted"] = False
            return True
        except pymssql.Error as e:
            logging.error(f"Error recovering task {task_id}", exc_info=True)
            messagebox.showerror("Database Error", f"Error recovering task: {str(e)}")
            return False

    def recover_all_deleted_tasks(self) -> int:
        """Recover all soft-deleted tasks."""
        recovered_count = 0
        
        try:
            # Update all deleted tasks in the database
            self.cursor.execute("UPDATE [dbo].[Tasks] SET deleted = 0 WHERE deleted = 1")
            recovered_count = self.cursor.rowcount
            self.conn.commit()
            
            # Update local copies
            for task in self._tasks:
                if task.get("deleted", False):
                    task["deleted"] = False
                    recovered_count += 1
            
            return recovered_count
        except pymssql.Error as e:
            logging.error(f"Error recovering all tasks: {str(e)}", exc_info=True)
            messagebox.showerror("Database Error", f"Error recovering tasks: {str(e)}")
            return 0

    def permanently_delete_task(self, task_id: int) -> bool:
        """Permanently remove a task from the system."""
        try:
            # Delete from the database
            self.cursor.execute("DELETE FROM [dbo].[Tasks] WHERE id = %s", (task_id,))
            self.conn.commit()
            
            # Remove from the local list
            self._tasks = [task for task in self._tasks if task["id"] != task_id]
            
            # Remove from deleted_tasks if it's there
            self._deleted_tasks = [t for t in self._deleted_tasks if t["id"] != task_id]
            
            return True
        except pymssql.Error as e:
            logging.error(f"Error permanently deleting task {task_id}", exc_info=True)
            messagebox.showerror("Database Error", f"Error permanently deleting task: {str(e)}")
            return False

    def permanently_delete_all_deleted_tasks(self) -> int:
        """Permanently remove all soft-deleted tasks."""
        try:
            # Delete all deleted tasks from the database
            self.cursor.execute("DELETE FROM [dbo].[Tasks] WHERE deleted = 1")
            deleted_count = self.cursor.rowcount
            self.conn.commit()
            
            # Update local list
            original_count = len(self._tasks)
            self._tasks = [task for task in self._tasks if not task.get("deleted", False)]
            
            # Clear the deleted_tasks list
            self._deleted_tasks = []
            
            return deleted_count
        except pymssql.Error as e:
            logging.error(f"Error permanently deleting all tasks: {str(e)}", exc_info=True)
            messagebox.showerror("Database Error", f"Error permanently deleting tasks: {str(e)}")
            return 0

    @property
    def tasks(self):
        """Get the list of tasks. Use property for consistent access."""
        return self._tasks

    # Add a new method to lock a task for editing
    def lock_task_for_editing(self, task_id: int, user: str) -> bool:
        """
        Lock a task for editing by a specific user.
        Returns True if lock was acquired, False otherwise.
        """
        self.begin_transaction()
        
        try:
            # Check if task is already locked
            self.cursor.execute("SELECT locked_by, lock_expiry FROM Tasks WHERE id = %s", (task_id,))
            result = self.cursor.fetchone()
            
            if not result:
                self.rollback_transaction()
                return False  # Task not found
                
            locked_by, lock_expiry = result
            
            current_time = datetime.datetime.now()
            
            # If task is locked by someone else and lock hasn't expired
            if locked_by and locked_by != user and lock_expiry and lock_expiry > current_time:
                self.rollback_transaction()
                return False  # Task is locked by another user
                
            # Set lock expiry to 5 minutes from now
            lock_expiry = current_time + datetime.timedelta(minutes=5)
            
            # Lock the task
            self.cursor.execute(
                "UPDATE Tasks SET locked_by = %s, lock_expiry = %s WHERE id = %s",
                (user, lock_expiry, task_id)
            )
            
            self.commit_transaction()
            return True
            
        except Exception as e:
            print(f"Error locking task: {e}")
            self.rollback_transaction()
            return False

    # Add a method to release a lock
    def release_task_lock(self, task_id: int, user: str) -> bool:
        """Release a lock on a task if it's locked by the specified user."""
        self.begin_transaction()
        
        try:
            # Check if task is locked by this user
            self.cursor.execute("SELECT locked_by FROM Tasks WHERE id = %s", (task_id,))
            result = self.cursor.fetchone()
            
            if not result:
                self.rollback_transaction()
                return False  # Task not found
                
            locked_by = result[0]
            
            # If task is locked by someone else
            if locked_by and locked_by != user:
                self.rollback_transaction()
                return False  # Task is locked by another user
                
            # Release the lock
            self.cursor.execute(
                "UPDATE Tasks SET locked_by = NULL, lock_expiry = NULL WHERE id = %s",
                (task_id,)
            )
            
            self.commit_transaction()
            return True
            
        except Exception as e:
            print(f"Error releasing task lock: {e}")
            self.rollback_transaction()
            return False

    # Make sure to add these columns to the database
    def _ensure_lock_fields_exist(self):
        """Ensure lock columns exist in the database."""
        try:
            # Check if locked_by column exists - using SQL Server syntax
            self.cursor.execute("""
                SELECT COUNT(*) 
                FROM INFORMATION_SCHEMA.COLUMNS 
                WHERE TABLE_NAME = 'Tasks' 
                AND COLUMN_NAME = 'locked_by'
            """)
            
            locked_by_exists = self.cursor.fetchone()[0] > 0
            
            if not locked_by_exists:
                # Add locked_by column
                self.cursor.execute("""
                    ALTER TABLE Tasks 
                    ADD locked_by NVARCHAR(255) NULL
                """)
                self.conn.commit()
                logging.info("Added 'locked_by' column to Tasks table")
                
            # Check if lock_expiry column exists
            self.cursor.execute("""
                SELECT COUNT(*) 
                FROM INFORMATION_SCHEMA.COLUMNS 
                WHERE TABLE_NAME = 'Tasks' 
                AND COLUMN_NAME = 'lock_expiry'
            """)
            
            lock_expiry_exists = self.cursor.fetchone()[0] > 0
            
            if not lock_expiry_exists:
                # Add lock_expiry column
                self.cursor.execute("""
                    ALTER TABLE Tasks 
                    ADD lock_expiry DATETIME NULL
                """)
                self.conn.commit()
                logging.info("Added 'lock_expiry' column to Tasks table")
                
        except pymssql.Error as e:
            logging.error(f"Error checking/adding lock columns: {str(e)}", exc_info=True)
            # We'll continue even if this fails

    def clean_expired_locks(self):
        """Clean up expired locks from the database."""
        try:
            current_time = datetime.datetime.now()
            self.cursor.execute(
                "UPDATE Tasks SET locked_by = NULL, lock_expiry = NULL WHERE lock_expiry IS NOT NULL AND lock_expiry < %s",
                (current_time,)
            )
            self.conn.commit()
            logging.info("Cleaned up expired locks")
        except Exception as e:
            logging.error(f"Error cleaning expired locks: {str(e)}", exc_info=True)


def main():
    """Main function to run the Task Manager application."""
    root = tk.Tk()
    task_manager = TaskManager() # Initialize TaskManager first - no arguments needed now
    app = TaskManagerApp(root, task_manager=task_manager)
    app.run()

if __name__ == "__main__":
    main()