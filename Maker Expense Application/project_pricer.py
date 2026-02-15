"""
Project Pricer - A tool for makers and DIYers to price out projects
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
from datetime import datetime
import json

class ProjectPricerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Project Pricer - Maker Edition")
        self.root.geometry("900x700")
        
        # Initialize database
        self.init_database()
        
        # Current selections
        self.current_profile_id = None
        self.current_project_id = None
        
        # Create UI
        self.create_menu()
        self.create_main_layout()
        
    def init_database(self):
        """Initialize SQLite database with required tables"""
        self.conn = sqlite3.connect('project_pricer.db')
        self.cursor = self.conn.cursor()
        
        # User Profile table
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS profiles (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                hourly_rate REAL NOT NULL,
                created_date TEXT
            )
        ''')
        
        # Tools/Machines table
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS tools (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                profile_id INTEGER,
                name TEXT NOT NULL,
                cost_per_hour REAL,
                FOREIGN KEY (profile_id) REFERENCES profiles (id)
            )
        ''')
        
        # Projects table
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                profile_id INTEGER,
                name TEXT NOT NULL,
                description TEXT,
                created_date TEXT,
                FOREIGN KEY (profile_id) REFERENCES profiles (id)
            )
        ''')
        
        # Materials table
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS materials (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER,
                name TEXT NOT NULL,
                quantity REAL,
                unit_cost REAL,
                FOREIGN KEY (project_id) REFERENCES projects (id)
            )
        ''')
        
        # Labor entries table
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS labor (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER,
                description TEXT,
                hours REAL,
                FOREIGN KEY (project_id) REFERENCES projects (id)
            )
        ''')
        
        # Tool usage table
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS tool_usage (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER,
                tool_id INTEGER,
                hours REAL,
                FOREIGN KEY (project_id) REFERENCES projects (id),
                FOREIGN KEY (tool_id) REFERENCES tools (id)
            )
        ''')
        
        self.conn.commit()
    
    def create_menu(self):
        """Create application menu bar"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New Profile", command=self.show_profile_dialog)
        file_menu.add_command(label="New Project", command=self.show_project_dialog)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # Export menu
        export_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Export", menu=export_menu)
        export_menu.add_command(label="Export to Excel", command=self.export_to_excel)
    
    def create_main_layout(self):
        """Create main application layout"""
        # Create notebook (tabbed interface)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Profile tab
        self.profile_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.profile_frame, text='Profile')
        self.create_profile_tab()
        
        # Projects tab
        self.projects_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.projects_frame, text='Projects')
        self.create_projects_tab()
        
        # Current Project tab
        self.current_project_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.current_project_frame, text='Current Project')
        self.create_current_project_tab()
    
    def create_profile_tab(self):
        """Create profile management tab"""
        # Profile selection
        selection_frame = ttk.LabelFrame(self.profile_frame, text="Select Profile", padding=10)
        selection_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(selection_frame, text="Profile:").grid(row=0, column=0, sticky='w')
        self.profile_var = tk.StringVar()
        self.profile_combo = ttk.Combobox(selection_frame, textvariable=self.profile_var, state='readonly', width=30)
        self.profile_combo.grid(row=0, column=1, padx=5)
        self.profile_combo.bind('<<ComboboxSelected>>', self.on_profile_selected)
        
        ttk.Button(selection_frame, text="New Profile", command=self.show_profile_dialog).grid(row=0, column=2, padx=5)
        
        # Profile details
        details_frame = ttk.LabelFrame(self.profile_frame, text="Profile Details", padding=10)
        details_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        ttk.Label(details_frame, text="Hourly Rate: $").grid(row=0, column=0, sticky='w')
        self.hourly_rate_label = ttk.Label(details_frame, text="--")
        self.hourly_rate_label.grid(row=0, column=1, sticky='w')
        
        # Tools/Machines section
        ttk.Label(details_frame, text="Tools & Machines:", font=('TkDefaultFont', 10, 'bold')).grid(row=1, column=0, columnspan=2, sticky='w', pady=(10, 5))
        
        # Tools listbox
        tools_scroll_frame = ttk.Frame(details_frame)
        tools_scroll_frame.grid(row=2, column=0, columnspan=2, sticky='nsew', pady=5)
        details_frame.rowconfigure(2, weight=1)
        details_frame.columnconfigure(1, weight=1)
        
        scrollbar = ttk.Scrollbar(tools_scroll_frame)
        scrollbar.pack(side='right', fill='y')
        
        self.tools_listbox = tk.Listbox(tools_scroll_frame, yscrollcommand=scrollbar.set, height=8)
        self.tools_listbox.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=self.tools_listbox.yview)
        
        # Tool buttons
        tool_buttons = ttk.Frame(details_frame)
        tool_buttons.grid(row=3, column=0, columnspan=2, pady=5)
        ttk.Button(tool_buttons, text="Add Tool", command=self.add_tool).pack(side='left', padx=2)
        ttk.Button(tool_buttons, text="Remove Tool", command=self.remove_tool).pack(side='left', padx=2)
        
        self.refresh_profiles()
    
    def create_projects_tab(self):
        """Create projects list tab"""
        # Projects list
        list_frame = ttk.LabelFrame(self.projects_frame, text="All Projects", padding=10)
        list_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Treeview for projects
        columns = ('Name', 'Description', 'Date', 'Total Cost')
        self.projects_tree = ttk.Treeview(list_frame, columns=columns, show='tree headings', height=15)
        
        self.projects_tree.heading('#0', text='ID')
        self.projects_tree.column('#0', width=50)
        
        for col in columns:
            self.projects_tree.heading(col, text=col)
            self.projects_tree.column(col, width=150)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.projects_tree.yview)
        self.projects_tree.configure(yscrollcommand=scrollbar.set)
        
        self.projects_tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Buttons
        button_frame = ttk.Frame(self.projects_frame)
        button_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(button_frame, text="New Project", command=self.show_project_dialog).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Open Project", command=self.open_selected_project).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Delete Project", command=self.delete_selected_project).pack(side='left', padx=5)
        
        self.refresh_projects_list()
    
    def create_current_project_tab(self):
        """Create current project editing tab"""
        # Project info
        info_frame = ttk.LabelFrame(self.current_project_frame, text="Project Information", padding=10)
        info_frame.pack(fill='x', padx=10, pady=10)
        
        self.project_name_label = ttk.Label(info_frame, text="No project selected", font=('TkDefaultFont', 12, 'bold'))
        self.project_name_label.pack()
        
        # Materials section
        materials_frame = ttk.LabelFrame(self.current_project_frame, text="Materials", padding=10)
        materials_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Materials treeview
        mat_columns = ('Name', 'Quantity', 'Unit Cost', 'Total')
        self.materials_tree = ttk.Treeview(materials_frame, columns=mat_columns, show='headings', height=6)
        
        for col in mat_columns:
            self.materials_tree.heading(col, text=col)
            self.materials_tree.column(col, width=120)
        
        self.materials_tree.pack(fill='both', expand=True)
        
        mat_buttons = ttk.Frame(materials_frame)
        mat_buttons.pack(fill='x', pady=5)
        ttk.Button(mat_buttons, text="Add Material", command=self.add_material).pack(side='left', padx=2)
        ttk.Button(mat_buttons, text="Remove Material", command=self.remove_material).pack(side='left', padx=2)
        
        # Labor section
        labor_frame = ttk.LabelFrame(self.current_project_frame, text="Labor", padding=10)
        labor_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        labor_columns = ('Description', 'Hours', 'Rate', 'Total')
        self.labor_tree = ttk.Treeview(labor_frame, columns=labor_columns, show='headings', height=4)
        
        for col in labor_columns:
            self.labor_tree.heading(col, text=col)
            self.labor_tree.column(col, width=120)
        
        self.labor_tree.pack(fill='both', expand=True)
        
        labor_buttons = ttk.Frame(labor_frame)
        labor_buttons.pack(fill='x', pady=5)
        ttk.Button(labor_buttons, text="Add Labor", command=self.add_labor).pack(side='left', padx=2)
        ttk.Button(labor_buttons, text="Remove Labor", command=self.remove_labor).pack(side='left', padx=2)
        
        # Tool usage section
        tool_usage_frame = ttk.LabelFrame(self.current_project_frame, text="Tool Usage", padding=10)
        tool_usage_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        tool_columns = ('Tool', 'Hours', 'Rate', 'Total')
        self.tool_usage_tree = ttk.Treeview(tool_usage_frame, columns=tool_columns, show='headings', height=4)
        
        for col in tool_columns:
            self.tool_usage_tree.heading(col, text=col)
            self.tool_usage_tree.column(col, width=120)
        
        self.tool_usage_tree.pack(fill='both', expand=True)
        
        tool_buttons = ttk.Frame(tool_usage_frame)
        tool_buttons.pack(fill='x', pady=5)
        ttk.Button(tool_buttons, text="Add Tool Usage", command=self.add_tool_usage).pack(side='left', padx=2)
        ttk.Button(tool_buttons, text="Remove Tool Usage", command=self.remove_tool_usage).pack(side='left', padx=2)
        
        # Total cost
        total_frame = ttk.Frame(self.current_project_frame)
        total_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(total_frame, text="Total Project Cost:", font=('TkDefaultFont', 12, 'bold')).pack(side='left')
        self.total_cost_label = ttk.Label(total_frame, text="$0.00", font=('TkDefaultFont', 14, 'bold'), foreground='green')
        self.total_cost_label.pack(side='left', padx=10)
    
    def show_profile_dialog(self):
        """Show dialog to create new profile"""
        dialog = tk.Toplevel(self.root)
        dialog.title("New Profile")
        dialog.geometry("400x250")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Profile Name:").grid(row=0, column=0, padx=10, pady=10, sticky='w')
        name_entry = ttk.Entry(dialog, width=30)
        name_entry.grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="Hourly Rate ($):").grid(row=1, column=0, padx=10, pady=10, sticky='w')
        rate_entry = ttk.Entry(dialog, width=30)
        rate_entry.grid(row=1, column=1, padx=10, pady=10)
        
        def save_profile():
            name = name_entry.get().strip()
            rate = rate_entry.get().strip()
            
            if not name or not rate:
                messagebox.showerror("Error", "Please fill in all fields")
                return
            
            try:
                rate = float(rate)
                self.cursor.execute(
                    'INSERT INTO profiles (name, hourly_rate, created_date) VALUES (?, ?, ?)',
                    (name, rate, datetime.now().isoformat())
                )
                self.conn.commit()
                messagebox.showinfo("Success", "Profile created successfully!")
                dialog.destroy()
                self.refresh_profiles()
            except ValueError:
                messagebox.showerror("Error", "Hourly rate must be a number")
        
        ttk.Button(dialog, text="Save", command=save_profile).grid(row=2, column=0, columnspan=2, pady=20)
    
    def show_project_dialog(self):
        """Show dialog to create new project"""
        if not self.current_profile_id:
            messagebox.showerror("Error", "Please select a profile first")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("New Project")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Project Name:").grid(row=0, column=0, padx=10, pady=10, sticky='w')
        name_entry = ttk.Entry(dialog, width=30)
        name_entry.grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="Description:").grid(row=1, column=0, padx=10, pady=10, sticky='nw')
        desc_text = tk.Text(dialog, width=30, height=6)
        desc_text.grid(row=1, column=1, padx=10, pady=10)
        
        def save_project():
            name = name_entry.get().strip()
            description = desc_text.get('1.0', 'end-1c').strip()
            
            if not name:
                messagebox.showerror("Error", "Please enter a project name")
                return
            
            self.cursor.execute(
                'INSERT INTO projects (profile_id, name, description, created_date) VALUES (?, ?, ?, ?)',
                (self.current_profile_id, name, description, datetime.now().isoformat())
            )
            self.conn.commit()
            messagebox.showinfo("Success", "Project created successfully!")
            dialog.destroy()
            self.refresh_projects_list()
        
        ttk.Button(dialog, text="Save", command=save_project).grid(row=2, column=0, columnspan=2, pady=20)
    
    def refresh_profiles(self):
        """Refresh profile dropdown"""
        self.cursor.execute('SELECT id, name FROM profiles')
        profiles = self.cursor.fetchall()
        
        profile_names = [f"{p[1]} (ID: {p[0]})" for p in profiles]
        self.profile_combo['values'] = profile_names
        
        if profiles and not self.current_profile_id:
            self.profile_combo.current(0)
            self.on_profile_selected(None)
    
    def on_profile_selected(self, event):
        """Handle profile selection"""
        selection = self.profile_var.get()
        if not selection:
            return
        
        # Extract ID from selection
        profile_id = int(selection.split('ID: ')[1].rstrip(')'))
        self.current_profile_id = profile_id
        
        # Load profile details
        self.cursor.execute('SELECT hourly_rate FROM profiles WHERE id = ?', (profile_id,))
        result = self.cursor.fetchone()
        
        if result:
            self.hourly_rate_label.config(text=f"{result[0]:.2f}")
        
        # Load tools
        self.refresh_tools()
    
    def refresh_tools(self):
        """Refresh tools list"""
        self.tools_listbox.delete(0, tk.END)
        
        if self.current_profile_id:
            self.cursor.execute('SELECT id, name, cost_per_hour FROM tools WHERE profile_id = ?', (self.current_profile_id,))
            tools = self.cursor.fetchall()
            
            for tool in tools:
                self.tools_listbox.insert(tk.END, f"{tool[1]} - ${tool[2]:.2f}/hr (ID: {tool[0]})")
    
    def add_tool(self):
        """Add a tool to the current profile"""
        if not self.current_profile_id:
            messagebox.showerror("Error", "Please select a profile first")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Tool")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Tool Name:").grid(row=0, column=0, padx=10, pady=10, sticky='w')
        name_entry = ttk.Entry(dialog, width=30)
        name_entry.grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="Cost per Hour ($):").grid(row=1, column=0, padx=10, pady=10, sticky='w')
        cost_entry = ttk.Entry(dialog, width=30)
        cost_entry.grid(row=1, column=1, padx=10, pady=10)
        
        def save_tool():
            name = name_entry.get().strip()
            cost = cost_entry.get().strip()
            
            if not name or not cost:
                messagebox.showerror("Error", "Please fill in all fields")
                return
            
            try:
                cost = float(cost)
                self.cursor.execute(
                    'INSERT INTO tools (profile_id, name, cost_per_hour) VALUES (?, ?, ?)',
                    (self.current_profile_id, name, cost)
                )
                self.conn.commit()
                messagebox.showinfo("Success", "Tool added successfully!")
                dialog.destroy()
                self.refresh_tools()
            except ValueError:
                messagebox.showerror("Error", "Cost must be a number")
        
        ttk.Button(dialog, text="Save", command=save_tool).grid(row=2, column=0, columnspan=2, pady=20)
    
    def remove_tool(self):
        """Remove selected tool"""
        selection = self.tools_listbox.curselection()
        if not selection:
            messagebox.showerror("Error", "Please select a tool to remove")
            return
        
        tool_text = self.tools_listbox.get(selection[0])
        tool_id = int(tool_text.split('ID: ')[1].rstrip(')'))
        
        if messagebox.askyesno("Confirm", "Are you sure you want to remove this tool?"):
            self.cursor.execute('DELETE FROM tools WHERE id = ?', (tool_id,))
            self.conn.commit()
            self.refresh_tools()
    
    def refresh_projects_list(self):
        """Refresh projects treeview"""
        # Clear existing items
        for item in self.projects_tree.get_children():
            self.projects_tree.delete(item)
        
        # Load all projects
        self.cursor.execute('''
            SELECT id, name, description, created_date 
            FROM projects 
            ORDER BY created_date DESC
        ''')
        projects = self.cursor.fetchall()
        
        for project in projects:
            # Calculate total cost
            total_cost = self.calculate_project_cost(project[0])
            
            self.projects_tree.insert('', 'end', text=str(project[0]),
                                    values=(project[1], project[2][:50], project[3][:10], f"${total_cost:.2f}"))
    
    def calculate_project_cost(self, project_id):
        """Calculate total cost for a project"""
        total = 0.0
        
        # Materials cost
        self.cursor.execute('SELECT quantity, unit_cost FROM materials WHERE project_id = ?', (project_id,))
        materials = self.cursor.fetchall()
        for mat in materials:
            total += (mat[0] or 0) * (mat[1] or 0)
        
        # Labor cost
        self.cursor.execute('''
            SELECT l.hours, p.hourly_rate 
            FROM labor l
            JOIN projects pr ON l.project_id = pr.id
            JOIN profiles p ON pr.profile_id = p.id
            WHERE l.project_id = ?
        ''', (project_id,))
        labor = self.cursor.fetchall()
        for lab in labor:
            total += (lab[0] or 0) * (lab[1] or 0)
        
        # Tool usage cost
        self.cursor.execute('''
            SELECT tu.hours, t.cost_per_hour
            FROM tool_usage tu
            JOIN tools t ON tu.tool_id = t.id
            WHERE tu.project_id = ?
        ''', (project_id,))
        tools = self.cursor.fetchall()
        for tool in tools:
            total += (tool[0] or 0) * (tool[1] or 0)
        
        return total
    
    def open_selected_project(self):
        """Open selected project for editing"""
        selection = self.projects_tree.selection()
        if not selection:
            messagebox.showerror("Error", "Please select a project to open")
            return
        
        project_id = int(self.projects_tree.item(selection[0], 'text'))
        self.current_project_id = project_id
        
        # Load project details
        self.cursor.execute('SELECT name, description FROM projects WHERE id = ?', (project_id,))
        project = self.cursor.fetchone()
        
        if project:
            self.project_name_label.config(text=project[0])
            self.refresh_current_project()
            self.notebook.select(self.current_project_frame)
    
    def delete_selected_project(self):
        """Delete selected project"""
        selection = self.projects_tree.selection()
        if not selection:
            messagebox.showerror("Error", "Please select a project to delete")
            return
        
        if messagebox.askyesno("Confirm", "Are you sure you want to delete this project?"):
            project_id = int(self.projects_tree.item(selection[0], 'text'))
            
            # Delete related records
            self.cursor.execute('DELETE FROM materials WHERE project_id = ?', (project_id,))
            self.cursor.execute('DELETE FROM labor WHERE project_id = ?', (project_id,))
            self.cursor.execute('DELETE FROM tool_usage WHERE project_id = ?', (project_id,))
            self.cursor.execute('DELETE FROM projects WHERE id = ?', (project_id,))
            self.conn.commit()
            
            self.refresh_projects_list()
            messagebox.showinfo("Success", "Project deleted successfully")
    
    def refresh_current_project(self):
        """Refresh current project view"""
        if not self.current_project_id:
            return
        
        # Clear all trees
        for item in self.materials_tree.get_children():
            self.materials_tree.delete(item)
        for item in self.labor_tree.get_children():
            self.labor_tree.delete(item)
        for item in self.tool_usage_tree.get_children():
            self.tool_usage_tree.delete(item)
        
        # Load materials
        self.cursor.execute('SELECT id, name, quantity, unit_cost FROM materials WHERE project_id = ?',
                          (self.current_project_id,))
        materials = self.cursor.fetchall()
        for mat in materials:
            total = (mat[2] or 0) * (mat[3] or 0)
            self.materials_tree.insert('', 'end', text=str(mat[0]),
                                      values=(mat[1], mat[2], f"${mat[3]:.2f}", f"${total:.2f}"))
        
        # Load labor
        self.cursor.execute('''
            SELECT l.id, l.description, l.hours, p.hourly_rate
            FROM labor l
            JOIN projects pr ON l.project_id = pr.id
            JOIN profiles p ON pr.profile_id = p.id
            WHERE l.project_id = ?
        ''', (self.current_project_id,))
        labor = self.cursor.fetchall()
        for lab in labor:
            total = (lab[2] or 0) * (lab[3] or 0)
            self.labor_tree.insert('', 'end', text=str(lab[0]),
                                  values=(lab[1], lab[2], f"${lab[3]:.2f}/hr", f"${total:.2f}"))
        
        # Load tool usage
        self.cursor.execute('''
            SELECT tu.id, t.name, tu.hours, t.cost_per_hour
            FROM tool_usage tu
            JOIN tools t ON tu.tool_id = t.id
            WHERE tu.project_id = ?
        ''', (self.current_project_id,))
        tools = self.cursor.fetchall()
        for tool in tools:
            total = (tool[2] or 0) * (tool[3] or 0)
            self.tool_usage_tree.insert('', 'end', text=str(tool[0]),
                                       values=(tool[1], tool[2], f"${tool[3]:.2f}/hr", f"${total:.2f}"))
        
        # Update total
        total_cost = self.calculate_project_cost(self.current_project_id)
        self.total_cost_label.config(text=f"${total_cost:.2f}")
    
    def add_material(self):
        """Add material to current project"""
        if not self.current_project_id:
            messagebox.showerror("Error", "No project selected")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Material")
        dialog.geometry("400x250")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Material Name:").grid(row=0, column=0, padx=10, pady=10, sticky='w')
        name_entry = ttk.Entry(dialog, width=30)
        name_entry.grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="Quantity:").grid(row=1, column=0, padx=10, pady=10, sticky='w')
        qty_entry = ttk.Entry(dialog, width=30)
        qty_entry.grid(row=1, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="Unit Cost ($):").grid(row=2, column=0, padx=10, pady=10, sticky='w')
        cost_entry = ttk.Entry(dialog, width=30)
        cost_entry.grid(row=2, column=1, padx=10, pady=10)
        
        def save_material():
            name = name_entry.get().strip()
            qty = qty_entry.get().strip()
            cost = cost_entry.get().strip()
            
            if not name or not qty or not cost:
                messagebox.showerror("Error", "Please fill in all fields")
                return
            
            try:
                qty = float(qty)
                cost = float(cost)
                self.cursor.execute(
                    'INSERT INTO materials (project_id, name, quantity, unit_cost) VALUES (?, ?, ?, ?)',
                    (self.current_project_id, name, qty, cost)
                )
                self.conn.commit()
                messagebox.showinfo("Success", "Material added successfully!")
                dialog.destroy()
                self.refresh_current_project()
            except ValueError:
                messagebox.showerror("Error", "Quantity and cost must be numbers")
        
        ttk.Button(dialog, text="Save", command=save_material).grid(row=3, column=0, columnspan=2, pady=20)
    
    def remove_material(self):
        """Remove selected material"""
        selection = self.materials_tree.selection()
        if not selection:
            messagebox.showerror("Error", "Please select a material to remove")
            return
        
        material_id = int(self.materials_tree.item(selection[0], 'text'))
        self.cursor.execute('DELETE FROM materials WHERE id = ?', (material_id,))
        self.conn.commit()
        self.refresh_current_project()
    
    def add_labor(self):
        """Add labor entry to current project"""
        if not self.current_project_id:
            messagebox.showerror("Error", "No project selected")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Labor")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Description:").grid(row=0, column=0, padx=10, pady=10, sticky='w')
        desc_entry = ttk.Entry(dialog, width=30)
        desc_entry.grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="Hours:").grid(row=1, column=0, padx=10, pady=10, sticky='w')
        hours_entry = ttk.Entry(dialog, width=30)
        hours_entry.grid(row=1, column=1, padx=10, pady=10)
        
        def save_labor():
            desc = desc_entry.get().strip()
            hours = hours_entry.get().strip()
            
            if not desc or not hours:
                messagebox.showerror("Error", "Please fill in all fields")
                return
            
            try:
                hours = float(hours)
                self.cursor.execute(
                    'INSERT INTO labor (project_id, description, hours) VALUES (?, ?, ?)',
                    (self.current_project_id, desc, hours)
                )
                self.conn.commit()
                messagebox.showinfo("Success", "Labor added successfully!")
                dialog.destroy()
                self.refresh_current_project()
            except ValueError:
                messagebox.showerror("Error", "Hours must be a number")
        
        ttk.Button(dialog, text="Save", command=save_labor).grid(row=2, column=0, columnspan=2, pady=20)
    
    def remove_labor(self):
        """Remove selected labor entry"""
        selection = self.labor_tree.selection()
        if not selection:
            messagebox.showerror("Error", "Please select a labor entry to remove")
            return
        
        labor_id = int(self.labor_tree.item(selection[0], 'text'))
        self.cursor.execute('DELETE FROM labor WHERE id = ?', (labor_id,))
        self.conn.commit()
        self.refresh_current_project()
    
    def add_tool_usage(self):
        """Add tool usage to current project"""
        if not self.current_project_id:
            messagebox.showerror("Error", "No project selected")
            return
        
        # Get available tools for current profile
        self.cursor.execute('''
            SELECT t.id, t.name, t.cost_per_hour
            FROM tools t
            JOIN projects p ON t.profile_id = p.profile_id
            WHERE p.id = ?
        ''', (self.current_project_id,))
        tools = self.cursor.fetchall()
        
        if not tools:
            messagebox.showerror("Error", "No tools available. Please add tools to your profile first.")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Tool Usage")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Tool:").grid(row=0, column=0, padx=10, pady=10, sticky='w')
        tool_var = tk.StringVar()
        tool_combo = ttk.Combobox(dialog, textvariable=tool_var, state='readonly', width=27)
        tool_combo['values'] = [f"{t[1]} (${t[2]:.2f}/hr)" for t in tools]
        tool_combo.grid(row=0, column=1, padx=10, pady=10)
        tool_combo.current(0)
        
        ttk.Label(dialog, text="Hours:").grid(row=1, column=0, padx=10, pady=10, sticky='w')
        hours_entry = ttk.Entry(dialog, width=30)
        hours_entry.grid(row=1, column=1, padx=10, pady=10)
        
        def save_tool_usage():
            hours = hours_entry.get().strip()
            
            if not hours:
                messagebox.showerror("Error", "Please enter hours")
                return
            
            try:
                hours = float(hours)
                tool_index = tool_combo.current()
                tool_id = tools[tool_index][0]
                
                self.cursor.execute(
                    'INSERT INTO tool_usage (project_id, tool_id, hours) VALUES (?, ?, ?)',
                    (self.current_project_id, tool_id, hours)
                )
                self.conn.commit()
                messagebox.showinfo("Success", "Tool usage added successfully!")
                dialog.destroy()
                self.refresh_current_project()
            except ValueError:
                messagebox.showerror("Error", "Hours must be a number")
        
        ttk.Button(dialog, text="Save", command=save_tool_usage).grid(row=2, column=0, columnspan=2, pady=20)
    
    def remove_tool_usage(self):
        """Remove selected tool usage"""
        selection = self.tool_usage_tree.selection()
        if not selection:
            messagebox.showerror("Error", "Please select a tool usage entry to remove")
            return
        
        usage_id = int(self.tool_usage_tree.item(selection[0], 'text'))
        self.cursor.execute('DELETE FROM tool_usage WHERE id = ?', (usage_id,))
        self.conn.commit()
        self.refresh_current_project()
    
    def export_to_excel(self):
        """Export current project to Excel"""
        if not self.current_project_id:
            messagebox.showerror("Error", "No project selected")
            return
        
        try:
            from excel_export import export_project_to_excel
            
            # Get project name for default filename
            self.cursor.execute('SELECT name FROM projects WHERE id = ?', (self.current_project_id,))
            project_name = self.cursor.fetchone()[0]
            
            # Ask user where to save
            default_filename = f"{project_name.replace(' ', '_')}_estimate.xlsx"
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                initialfile=default_filename
            )
            
            if filename:
                export_project_to_excel(self.cursor, self.current_project_id, filename)
                messagebox.showinfo("Success", f"Project exported to:\n{filename}")
                
        except ImportError:
            messagebox.showerror("Missing Dependency", 
                              "Excel export requires the 'openpyxl' library.\n\n"
                              "Install it with: pip install openpyxl\n\n"
                              "Then restart the application.")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export project:\n{str(e)}")
    
    def __del__(self):
        """Close database connection"""
        if hasattr(self, 'conn'):
            self.conn.close()

def main():
    root = tk.Tk()
    app = ProjectPricerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()