import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import json
import os
import sys
import webbrowser
from io import StringIO
from datetime import datetime
from typing import Optional, List, Dict, Any
import queue
import time
import asyncio
from concurrent.futures import ThreadPoolExecutor
import re

# Import existing agent dependencies
try:
    from langchain_google_genai import ChatGoogleGenerativeAI
    from windows_use.agent import Agent
    from dotenv import load_dotenv
    DEPENDENCIES_AVAILABLE = True
except ImportError:
    DEPENDENCIES_AVAILABLE = False

# Load environment variables
load_dotenv()

class WindowsAgentGUI:
    """
    Modern GUI application for Windows Agent using tkinter.
    Provides a professional interface while preserving all original agent functionality.
    """
    
    def __init__(self):
        # Main window setup
        self.root = tk.Tk()
        self.root.title("Windows Agent - AI Desktop Assistant")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 700)
        
        # Configure dark theme colors
        self.colors = {
            'bg': '#f8f9fa',
            'fg': '#333333',
            'sidebar_bg': '#ffffff',
            'button_bg': '#f1f3f4',
            'entry_bg': '#ffffff',
            'text_bg': '#ffffff',
            'green': '#28a745',
            'red': '#dc3545',
            'gray': '#6c757d',
            'blue': '#007bff',
            'selected_bg': '#e8e8e8',
            'card_bg': '#ffffff',
            'border': '#e0e0e0',
            # Hover colors for buttons
            'green_hover': '#218838',
            'red_hover': '#c82333',
            'gray_hover': '#5a6268',
            'blue_hover': '#0056b3'
        }
        
        # Configure root window
        self.root.configure(bg=self.colors['bg'])
        
        # Application state
        self.agent: Optional[Agent] = None
        self.llm: Optional[ChatGoogleGenerativeAI] = None
        self.current_api_key: str = ""  # Store API key in memory
        self.is_task_running = False
        self.current_task: Optional[asyncio.Task] = None
        self.task_thread: Optional[threading.Thread] = None
        self.message_queue = queue.Queue()
        
        # Chat history and settings
        self.chat_history: List[Dict[str, Any]] = []
        self.task_history: List[str] = []
        self.settings = self.load_settings()
        
        # Agent output capture
        self.original_stdout = sys.stdout
        self.original_stderr = sys.stderr
        
        # GUI components
        self.setup_gui()
        self.load_chat_history()
        self.load_task_history()
        
        # Initialize agent if dependencies are available
        if DEPENDENCIES_AVAILABLE:
            self.load_api_key_from_env()
            # Load rules from .env on startup
            self.settings["rules"] = self.load_rules()
        
        # Start message processing
        self.process_message_queue()
        
        # Handle window closing
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def setup_gui(self):
        """
        Create and configure all GUI components with proper layout and styling.
        """
        # Configure grid weights for responsive design
        self.root.grid_columnconfigure(1, weight=3)  # Main content area gets more space
        self.root.grid_columnconfigure(0, weight=1)  # Sidebar gets less space
        self.root.grid_rowconfigure(0, weight=1)
        
        # Create main containers
        self.create_sidebar()
        self.create_main_panel()
        self.create_settings_panel()
    
    def create_sidebar(self):
        """
        Create left sidebar with edit icon, history, and control buttons.
        """
        # Sidebar frame
        self.sidebar_frame = tk.Frame(self.root, bg=self.colors['sidebar_bg'], width=280, relief='solid', borderwidth=1)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 2))
        self.sidebar_frame.grid_rowconfigure(2, weight=1)  # Task history expands
        self.sidebar_frame.grid_columnconfigure(0, weight=1)
        
        # Header with edit icon
        header_frame = tk.Frame(self.sidebar_frame, bg=self.colors['sidebar_bg'])
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        
        edit_icon = tk.Label(
            header_frame,
            text="✏️",
            font=("Arial", 16),
            bg=self.colors['sidebar_bg'],
            fg=self.colors['fg'],
            cursor='hand2'
        )
        edit_icon.pack(anchor="w")
        
        # History section
        history_label = tk.Label(
            self.sidebar_frame, 
            text="HISTORY", 
            font=("Arial", 10, "bold"),
            bg=self.colors['sidebar_bg'],
            fg=self.colors['gray']
        )
        history_label.grid(row=1, column=0, padx=20, pady=(20, 5), sticky="w")
        
        # Task history container
        history_container = tk.Frame(self.sidebar_frame, bg=self.colors['sidebar_bg'])
        history_container.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="nsew")
        history_container.grid_rowconfigure(0, weight=1)
        history_container.grid_columnconfigure(0, weight=1)
        
        # Scrollable task history
        self.task_canvas = tk.Canvas(history_container, bg=self.colors['text_bg'], highlightthickness=0)
        self.task_scrollbar = ttk.Scrollbar(history_container, orient="vertical", command=self.task_canvas.yview)
        self.task_history_frame = tk.Frame(self.task_canvas, bg=self.colors['text_bg'])
        
        self.task_canvas.configure(yscrollcommand=self.task_scrollbar.set)
        self.task_canvas.grid(row=0, column=0, sticky="nsew")
        self.task_scrollbar.grid(row=0, column=1, sticky="ns")
        
        self.task_canvas_window = self.task_canvas.create_window((0, 0), window=self.task_history_frame, anchor="nw")
        self.task_history_frame.bind("<Configure>", self.on_task_frame_configure)
        self.task_canvas.bind("<Configure>", self.on_task_canvas_configure)
        
        # Store selected task for highlighting
        self.selected_task_widget = None
        
        # Control buttons at bottom
        buttons_frame = tk.Frame(self.sidebar_frame, bg=self.colors['sidebar_bg'])
        buttons_frame.grid(row=3, column=0, sticky="ew", padx=20, pady=(0, 20))
        buttons_frame.grid_columnconfigure(0, weight=1)
        
        # Start Task button
        self.start_button = tk.Button(
            buttons_frame,
            text="▶ Start Task",
            bg=self.colors['green'],
            fg='white',
            font=("Arial", 11, "bold"),
            command=self.start_task,
            height=1,
            relief='flat',
            cursor='hand2',
            pady=8
        )
        self.start_button.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        
        # Add hover effects for start button
        self.start_button.bind("<Enter>", lambda e: self._on_button_enter(self.start_button, self.colors['green_hover']))
        self.start_button.bind("<Leave>", lambda e: self._on_button_leave(self.start_button, self.colors['green']))
        
        # Stop Task button
        self.stop_button = tk.Button(
            buttons_frame,
            text="⏹ Stop Task",
            bg=self.colors['red'],
            fg='white',
            font=("Arial", 11, "bold"),
            command=self.stop_task,
            height=1,
            relief='flat',
            cursor='hand2',
            state="disabled",
            pady=8
        )
        self.stop_button.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        
        # Add hover effects for stop button
        self.stop_button.bind("<Enter>", lambda e: self._on_button_enter(self.stop_button, self.colors['red_hover']))
        self.stop_button.bind("<Leave>", lambda e: self._on_button_leave(self.stop_button, self.colors['red']))
        
        # Settings button
        self.settings_button = tk.Button(
            buttons_frame,
            text="⚙ Settings",
            bg=self.colors['gray'],
            fg='white',
            font=("Arial", 11, "bold"),
            command=self.show_settings,
            height=1,
            relief='flat',
            cursor='hand2',
            pady=8
        )
        self.settings_button.grid(row=2, column=0, sticky="ew")
        
        # Add hover effects for settings button
        self.settings_button.bind("<Enter>", lambda e: self._on_button_enter(self.settings_button, self.colors['gray_hover']))
        self.settings_button.bind("<Leave>", lambda e: self._on_button_leave(self.settings_button, self.colors['gray']))
    
    def on_task_frame_configure(self, event):
        """Update scroll region when task frame size changes."""
        self.task_canvas.configure(scrollregion=self.task_canvas.bbox("all"))
    
    def on_task_canvas_configure(self, event):
        """Update canvas window width when canvas size changes."""
        canvas_width = event.width
        self.task_canvas.itemconfig(self.task_canvas_window, width=canvas_width)
    
    def create_main_panel(self):
        """
        Create main content area with warning icon, command cards, and input field.
        """
        # Main panel frame
        self.main_frame = tk.Frame(self.root, bg=self.colors['bg'])
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=(2, 0))
        self.main_frame.grid_rowconfigure(1, weight=1)  # Command cards area expands
        self.main_frame.grid_columnconfigure(0, weight=1)
        
        # Command cards container
        cards_container = tk.Frame(self.main_frame, bg=self.colors['bg'])
        cards_container.grid(row=0, column=0, padx=40, pady=(60, 40), sticky="nsew")
        cards_container.grid_columnconfigure(0, weight=1)
        cards_container.grid_columnconfigure(1, weight=1)
        cards_container.grid_columnconfigure(2, weight=1)
        
        # Create three command cards
        self.create_command_card(cards_container, 0, 0, "🔵", "Commands", 
                               "Go to youtube and play song beat it.")
        self.create_command_card(cards_container, 1, 0, "🟡", "Commands", 
                               "Go to microsoft word and write a short 20 word summary on agi.")
        self.create_command_card(cards_container, 2, 0, "🟠", "Commands", 
                               "Go to microsoft copilot and ask what is agi.")
        
        # Input section at bottom
        input_frame = tk.Frame(self.main_frame, bg=self.colors['bg'])
        input_frame.grid(row=1, column=0, padx=40, pady=(0, 40), sticky="ew")
        input_frame.grid_columnconfigure(0, weight=1)
        
        # Command input field with updated styling
        self.command_entry = tk.Entry(
            input_frame,
            font=("Arial", 14),
            bg=self.colors['entry_bg'],
            fg=self.colors['fg'],
            insertbackground=self.colors['fg'],
            relief='solid',
            borderwidth=1,
            highlightthickness=0
        )
        self.command_entry.grid(row=0, column=0, sticky="ew", ipady=12, padx=2, pady=2)
        self.command_entry.bind("<Return>", self.send_command)
        
        # Updated placeholder text
        self.command_entry.insert(0, "Message Luma")
        self.command_entry.bind("<FocusIn>", self.clear_placeholder)
        self.command_entry.bind("<FocusOut>", self.add_placeholder)
        self.command_entry.config(fg='gray')
        
        # Disclaimer text
        disclaimer_label = tk.Label(
            input_frame,
            text="AI agents can make mistakes. Check important info.",
            font=("Arial", 10),
            bg=self.colors['bg'],
            fg=self.colors['gray']
        )
        disclaimer_label.grid(row=1, column=0, pady=(8, 0))
        
        # Main footer frame for branding
        main_footer_frame = tk.Frame(self.main_frame, bg=self.colors['bg'])
        main_footer_frame.grid(row=2, column=0, padx=40, pady=(20, 20), sticky="ew")
        main_footer_frame.grid_columnconfigure(0, weight=1)
        
        # Version label
        version_label = tk.Label(
            main_footer_frame,
            text="Luma – Windows Agent (Version 1.0)",
            font=("Arial", 11, "bold"),
            bg=self.colors['bg'],
            fg=self.colors['fg']
        )
        version_label.pack(pady=(0, 5))
        
        # Developer info frame to hold text and link inline
        dev_info_frame = tk.Frame(main_footer_frame, bg=self.colors['bg'])
        dev_info_frame.pack()
        
        # Developer text label
        dev_text_label = tk.Label(
            dev_info_frame,
            text="Luma is developed by Machovotrish. Learn more at ",
            font=("Arial", 10),
            bg=self.colors['bg'],
            fg=self.colors['gray']
        )
        dev_text_label.pack(side="left")
        
        # Clickable link label
        link_label = tk.Label(
            dev_info_frame,
            text="www.machovotrish.com",
            font=("Arial", 10, "underline"),
            bg=self.colors['bg'],
            fg=self.colors['blue'],
            cursor='hand2'
        )
        link_label.pack(side="left")
        
        # Bind click event to link
        link_label.bind("<Button-1>", self.open_machovotrish_link)
        
        # Add hover effects for the link
        link_label.bind("<Enter>", lambda e: link_label.configure(fg=self.colors['blue_hover']))
        link_label.bind("<Leave>", lambda e: link_label.configure(fg=self.colors['blue']))
        
        # Create chat display for showing agent logs and messages
        self.chat_display = scrolledtext.ScrolledText(
            input_frame,
            wrap=tk.WORD,
            state=tk.DISABLED,
            font=("Consolas", 10),
            bg=self.colors['text_bg'],
            fg=self.colors['fg'],
            height=15
        )
        self.chat_display.grid(row=3, column=0, sticky="ew", pady=(10, 0))
    
    def create_command_card(self, parent, column, row, icon, title, description):
        """
        Create a command suggestion card with icon, title, and description.
        """
        # Card frame
        card_frame = tk.Frame(
            parent,
            bg=self.colors['card_bg'],
            relief='solid',
            borderwidth=1,
            cursor='hand2'
        )
        card_frame.grid(row=row, column=column, padx=10, pady=10, sticky="nsew", ipadx=20, ipady=20)
        
        # Icon
        icon_label = tk.Label(
            card_frame,
            text=icon,
            font=("Arial", 24),
            bg=self.colors['card_bg'],
            fg=self.colors['fg']
        )
        icon_label.pack(pady=(0, 10))
        
        # Title
        title_label = tk.Label(
            card_frame,
            text=title,
            font=("Arial", 14, "bold"),
            bg=self.colors['card_bg'],
            fg=self.colors['fg']
        )
        title_label.pack(pady=(0, 8))
        
        # Description
        desc_label = tk.Label(
            card_frame,
            text=description,
            font=("Arial", 10),
            bg=self.colors['card_bg'],
            fg=self.colors['gray'],
            wraplength=200,
            justify='center'
        )
        desc_label.pack()
        
        # Bind click events to all card elements
        def on_card_click(event=None):
            self.load_previous_task(description)
        
        card_frame.bind("<Button-1>", on_card_click)
        icon_label.bind("<Button-1>", on_card_click)
        title_label.bind("<Button-1>", on_card_click)
        desc_label.bind("<Button-1>", on_card_click)
        
        # Hover effects
        def on_enter(event):
            # Enhanced shadow pop effect
            card_frame.configure(bg=self.colors['selected_bg'], relief='raised', borderwidth=2)
            icon_label.configure(bg=self.colors['selected_bg'])
            title_label.configure(bg=self.colors['selected_bg'])
            desc_label.configure(bg=self.colors['selected_bg'])
        
        def on_leave(event):
            # Restore original appearance
            card_frame.configure(bg=self.colors['card_bg'], relief='solid', borderwidth=1)
            icon_label.configure(bg=self.colors['card_bg'])
            title_label.configure(bg=self.colors['card_bg'])
            desc_label.configure(bg=self.colors['card_bg'])
        
        card_frame.bind("<Enter>", on_enter)
        card_frame.bind("<Leave>", on_leave)
    
    def clear_placeholder(self, event):
        """Clear placeholder text when entry gains focus."""
        if self.command_entry.get() == "Message Luma":
            self.command_entry.delete(0, tk.END)
            self.command_entry.config(fg=self.colors['fg'])
    
    def add_placeholder(self, event):
        """Add placeholder text when entry loses focus and is empty."""
        if not self.command_entry.get():
            self.command_entry.insert(0, "Message Luma")
            self.command_entry.config(fg='gray')
    
    def create_settings_panel(self):
        """
        Create settings panel as a separate window for configuration management.
        """
        self.settings_window = None  # Will be created when needed
    
    def show_settings(self):
        """
        Display settings configuration window with API key and rules management.
        """
        if self.settings_window is not None and self.settings_window.winfo_exists():
            self.settings_window.focus()
            return
        
        # Create settings window
        self.settings_window = tk.Toplevel(self.root)
        self.settings_window.title("Agent Settings")
        self.settings_window.geometry("800x600")
        self.settings_window.configure(bg=self.colors['bg'])
        self.settings_window.transient(self.root)
        self.settings_window.grab_set()
        
        # API Key section
        api_frame = tk.LabelFrame(
            self.settings_window, 
            text="API Configuration", 
            bg=self.colors['bg'],
            fg=self.colors['fg'],
            font=("Arial", 10, "bold")
        )
        api_frame.pack(fill="x", padx=20, pady=20)
        
        tk.Label(api_frame, text="Google API Key:", bg=self.colors['bg'], fg=self.colors['fg']).pack(anchor="w", padx=10, pady=(10, 5))
        
        # API key input with show/hide toggle
        key_frame = tk.Frame(api_frame, bg=self.colors['bg'])
        key_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        self.api_key_entry = tk.Entry(key_frame, show="*", bg=self.colors['entry_bg'], fg=self.colors['fg'], insertbackground=self.colors['fg'])
        self.api_key_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        self.show_key_var = tk.BooleanVar()
        self.show_key_checkbox = tk.Checkbutton(
            key_frame, 
            text="Show", 
            variable=self.show_key_var,
            command=self.toggle_api_key_visibility,
            bg=self.colors['bg'],
            fg=self.colors['fg'],
            selectcolor=self.colors['button_bg']
        )
        self.show_key_checkbox.pack(side="right")
        
        # Save API Key button
        self.save_api_key_button = tk.Button(
            key_frame,
            text="Save API Key",
            bg=self.colors['blue'],
            fg='white',
            font=("Arial", 9, "bold"),
            command=self.save_api_key,
            relief='flat',
            cursor='hand2'
        )
        self.save_api_key_button.pack(side="right", padx=(5, 0))
        
        # Add hover effects for save API key button
        self.save_api_key_button.bind("<Enter>", lambda e: self._on_button_enter(self.save_api_key_button, self.colors['blue_hover']))
        self.save_api_key_button.bind("<Leave>", lambda e: self._on_button_leave(self.save_api_key_button, self.colors['blue']))
        
        # Load current API key
        if self.current_api_key:
            self.api_key_entry.insert(0, self.current_api_key)
        
        # Rules section (Limited to 5 Rules)
        rules_frame = tk.LabelFrame(
            self.settings_window, 
            text="Agent Rules", 
            bg=self.colors['bg'],
            fg=self.colors['fg'],
            font=("Arial", 10, "bold")
        )
        rules_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        # Create 5 rule entry fields
        self.rule_entries = []
        for i in range(5):
            rule_container = tk.Frame(rules_frame, bg=self.colors['bg'])
            rule_container.pack(fill="x", padx=10, pady=5)
            
            rule_label = tk.Label(
                rule_container,
                text=f"Rule {i+1}:",
                bg=self.colors['bg'],
                fg=self.colors['fg'],
                font=("Arial", 10),
                width=8,
                anchor="w"
            )
            rule_label.pack(side="left", padx=(0, 10))
            
            rule_entry = tk.Entry(
                rule_container,
                bg=self.colors['entry_bg'],
                fg=self.colors['fg'],
                insertbackground=self.colors['fg'],
                font=("Arial", 10)
            )
            rule_entry.pack(side="left", fill="x", expand=True)
            self.rule_entries.append(rule_entry)
        
        # Load current rules into entry fields
        current_rules = self.load_rules()
        for i, rule in enumerate(current_rules[:5]):  # Only load up to 5 rules
            if i < len(self.rule_entries):
                self.rule_entries[i].insert(0, rule)
        
        # Save Rules button
        save_rules_frame = tk.Frame(rules_frame, bg=self.colors['bg'])
        save_rules_frame.pack(fill="x", padx=10, pady=(10, 10))
        
        self.save_rules_button = tk.Button(
            save_rules_frame,
            text="Save Rules",
            bg=self.colors['green'],
            fg='white',
            font=("Arial", 10, "bold"),
            command=self.save_rules_from_settings_ui,
            relief='flat',
            cursor='hand2'
        )
        self.save_rules_button.pack(side="right")
        
        # Add hover effects for save rules button
        self.save_rules_button.bind("<Enter>", lambda e: self._on_button_enter(self.save_rules_button, self.colors['green_hover']))
        self.save_rules_button.bind("<Leave>", lambda e: self._on_button_leave(self.save_rules_button, self.colors['green']))
        
        # Buttons frame
        buttons_frame = tk.Frame(self.settings_window, bg=self.colors['bg'])
        buttons_frame.pack(fill="x", padx=20, pady=(0, 20))
        
        tk.Button(
            buttons_frame, 
            text="Cancel", 
            bg=self.colors['gray'],
            fg='white',
            command=self.settings_window.destroy,
            relief='flat',
            cursor='hand2'
        ).pack(side="right", padx=(5, 0))
    
    def toggle_api_key_visibility(self):
        """Toggle API key visibility in settings panel."""
        if self.show_key_var.get():
            self.api_key_entry.configure(show="")
        else:
            self.api_key_entry.configure(show="*")
    
    def _on_button_enter(self, button, hover_color):
        """Handle button hover enter event with color transition."""
        if button['state'] != 'disabled':
            button.configure(bg=hover_color)
    
    def _on_button_leave(self, button, normal_color):
        """Handle button hover leave event with color restoration."""
        if button['state'] != 'disabled':
            button.configure(bg=normal_color)
    
    def save_api_key(self):
        """Save API key to .env file and reinitialize agent if needed."""
        try:
            # Get API key from input
            api_key = self.api_key_entry.get().strip()
            
            # Save API key to .env file and memory
            if api_key:
                self.save_api_key_to_env(api_key)
                self.current_api_key = api_key
                self.add_message("System", "API key saved successfully.", "system")
            else:
                self.add_message("System", "No API key provided.", "system")
            
            # Reinitialize agent with new settings
            if DEPENDENCIES_AVAILABLE and self.current_api_key:
                self.initialize_agent()
            else:
                self.add_message("System", "Agent will be initialized when you start a task.", "system")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save API key: {str(e)}")
    
    def save_rules_to_env(self):
        """Save rules to .env file."""
        try:
            # Get rules from entry fields
            rules = []
            for entry in self.rule_entries:
                rule = entry.get().strip()
                if rule and rule not in rules:  # Skip empty fields and prevent duplicates
                    rules.append(rule)
            
            # Save rules to .env
            self.save_rules(rules)
            
            # Update settings in memory
            self.settings["rules"] = rules
            
            self.add_message("System", f"Saved {len(rules)} rules successfully.", "system")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save rules: {str(e)}")
    
    def save_rules_from_settings_ui(self):
        """Save rules from settings UI to rules.json file."""
        try:
            # Get rules from entry fields
            rules = []
            for entry in self.rule_entries:
                rule = entry.get().strip()
                if rule and rule not in rules:  # Skip empty fields and prevent duplicates
                    rules.append(rule)
            
            # Save rules to JSON file
            self._save_rules_to_file(rules)
            
            # Update settings in memory
            self.settings["rules"] = rules
            
            self.add_message("System", f"Saved {len(rules)} rules successfully.", "system")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save rules: {str(e)}")
    
    def load_api_key_from_env(self):
        """Load API key from .env file if it exists."""
        try:
            # Load environment variables from .env file
            load_dotenv()
            
            # Get API key from environment
            api_key = os.getenv("GOOGLE_API_KEY")
            if api_key:
                self.current_api_key = api_key.strip()
            else:
                self.current_api_key = ""
                
        except Exception as e:
            print(f"Error loading API key from .env: {e}")
            self.current_api_key = ""
    
    def load_rules(self) -> List[str]:
        """Load rules from rules.json file."""
        rules = []
        try:
            if os.path.exists('rules.json'):
                with open('rules.json', 'r') as f:
                    rules = json.load(f)
                    # Ensure we have a list and filter out empty rules
                    if isinstance(rules, list):
                        rules = [rule.strip() for rule in rules if rule.strip()]
                    else:
                        rules = []
        except (FileNotFoundError, json.JSONDecodeError) as e:
            print(f"Error loading rules from rules.json: {e}")
            rules = []
        except Exception as e:
            print(f"Error loading rules from rules.json: {e}")
            rules = []
        
        return rules
    
    def _save_rules_to_file(self, rules_list: List[str]):
        """Save rules to rules.json file."""
        try:
            # Save rules as JSON array
            with open('rules.json', 'w') as f:
                json.dump(rules_list, f, indent=2)
                
        except Exception as e:
            print(f"Error saving rules to rules.json: {e}")
            raise
    
    def save_api_key_to_env(self, api_key: str):
        """Save API key to .env file using python-dotenv."""
        try:
            from dotenv import set_key
            
            # Use set_key to save API key to .env file
            set_key('.env', 'GOOGLE_API_KEY', api_key)
            
            # Reload environment variables to make the new key available immediately
            load_dotenv(override=True)
                
        except Exception as e:
            print(f"Error saving API key to .env: {e}")
            raise
    
    def initialize_agent(self):
        """
        Initialize the agent with current settings (preserves original logic).
        """
        try:
            if not DEPENDENCIES_AVAILABLE:
                self.add_message("System", "Agent dependencies not available. Please install required packages.", "error")
                return
            
            # Initialize LLM (original logic preserved)
            self.llm = ChatGoogleGenerativeAI(model='gemini-2.0-flash')
            
            # Initialize Agent (original logic preserved)
            self.agent = Agent(llm=self.llm, browser='chrome', use_vision=True)
            
            self.add_message("System", "Agent initialized successfully.", "system")
            
        except Exception as e:
            self.add_message("System", f"Failed to initialize agent: {str(e)}", "error")
    
    def capture_agent_output(self, query: str):
        """
        Capture agent output and send to GUI while preserving terminal output.
        """
        # Create string buffer to capture output
        captured_output = StringIO()
        
        class TeeOutput:
            def __init__(self, original, captured, message_queue, sender="Agent"):
                self.original = original
                self.captured = captured
                self.message_queue = message_queue
                self.sender = sender
                self.buffer = ""
            
            def write(self, text):
                # Write to original (terminal)
                self.original.write(text)
                self.original.flush()
                
                # Write to captured buffer
                self.captured.write(text)
                
                # Buffer text and send complete lines to GUI
                self.buffer += text
                while '\n' in self.buffer:
                    line, self.buffer = self.buffer.split('\n', 1)
                    if line.strip():  # Only send non-empty lines
                        # Format agent logs for better readability
                        formatted_line = self.format_agent_log(line.strip())
                        if formatted_line:
                            self.message_queue.put((self.sender, formatted_line, "agent"))
            
            def flush(self):
                self.original.flush()
                self.captured.flush()
                # Send any remaining buffer content
                if self.buffer.strip():
                    formatted_line = self.format_agent_log(self.buffer.strip())
                    if formatted_line:
                        self.message_queue.put((self.sender, formatted_line, "agent"))
                    self.buffer = ""
            
            def format_agent_log(self, line):
                """Format agent log lines for better GUI display."""
                line = line.strip()
                if not line:
                    return None
                
                # Add emojis and formatting for different log types
                if "Iteration" in line:
                    return f"🔄 {line}"
                elif "Evaluate" in line:
                    return f"🔍 {line}"
                elif "Memory" in line:
                    return f"🧠 {line}"
                elif "Thought" in line:
                    return f"💭 {line}"
                elif "Action" in line:
                    return f"⚡ {line}"
                elif "Observation" in line:
                    return f"👁️ {line}"
                elif "Final Answer" in line:
                    return f"✅ {line}"
                elif "ERROR" in line.upper() or "EXCEPTION" in line.upper():
                    return f"❌ {line}"
                elif "WARNING" in line.upper():
                    return f"⚠️ {line}"
                else:
                    # Return other agent output as-is
                    return line
        
        # Set up output capture
        tee_stdout = TeeOutput(self.original_stdout, captured_output, self.message_queue, "Agent")
        tee_stderr = TeeOutput(self.original_stderr, captured_output, self.message_queue, "Agent")
        
        try:
            # Redirect stdout and stderr to capture agent output
            sys.stdout = tee_stdout
            sys.stderr = tee_stderr
            
            # Execute the agent task
            agent_result = self.agent.invoke(query)
            
            # Ensure any remaining output is flushed
            sys.stdout.flush()
            sys.stderr.flush()
            
            return agent_result
            
        finally:
            # Always restore original stdout/stderr
            sys.stdout = self.original_stdout
            sys.stderr = self.original_stderr
    
    def add_message(self, sender: str, message: str, msg_type: str = "user"):
        """
        Add message to chat display with proper formatting and auto-scroll.
        """
        # Enable text widget for editing
        self.chat_display.configure(state=tk.NORMAL)
        
        # Format timestamp
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # Insert formatted message
        self.chat_display.insert(tk.END, f"[{timestamp}] {sender}: {message}\n\n")
        
        # Auto-scroll to bottom
        self.chat_display.see(tk.END)
        
        # Disable editing
        self.chat_display.configure(state=tk.DISABLED)
        
        # Save to chat history
        self.chat_history.append({
            "timestamp": timestamp,
            "sender": sender,
            "message": message,
            "type": msg_type
        })
        
        # Save chat history to file
        self.save_chat_history()
    
    def send_command(self, event=None):
        """
        Process user command input and execute agent task.
        """
        command = self.command_entry.get().strip()
        if not command or command == "Message Luma":
            return
        
        # CHANGE: Add API key validation before proceeding
        if not self.current_api_key:
            self.add_message("System", "Error: Please set your Google API key in Settings before starting a task.", "error")
            # Automatically open Settings window
            self.show_settings()
            return
        
        # Clear input field
        self.command_entry.delete(0, tk.END)
        self.add_placeholder(None)
        
        # Add user message to chat
        self.add_message("User", command, "user")
        
        # Check if agent is available
        if not DEPENDENCIES_AVAILABLE:
            self.add_message("System", "Please install required dependencies: langchain-google-genai, windows-use", "error")
            return
        
        # Add to task history
        self.task_history.append(command)
        self.update_task_history_display()
        
        # Execute command in separate thread to prevent GUI blocking
        if not self.is_task_running:
            # CHANGE: Use asyncio for better task cancellation support
            self.task_thread = threading.Thread(target=self.run_agent_task_async, args=(command,))
            self.task_thread.daemon = True
            self.task_thread.start()
    
    def run_agent_task_async(self, query: str):
        """
        Run agent task in asyncio event loop for better cancellation support.
        """
        # Create new event loop for this thread
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        
        try:
            # CHANGE: Create cancellable task
            self.current_task = loop.create_task(self.execute_agent_task_async(query))
            loop.run_until_complete(self.current_task)
        except asyncio.CancelledError:
            self.message_queue.put(("System", "Task was cancelled by user.", "system"))
        except Exception as e:
            self.message_queue.put(("System", f"Error in task execution: {str(e)}", "error"))
        finally:
            loop.close()
            self.current_task = None
            # Add confirmation message after thread cleanup
            if not self.is_task_running:
                self.message_queue.put(("System", "✋ Task stopped successfully.", "system"))
    
    async def execute_agent_task_async(self, query: str):
        """
        Execute agent task asynchronously with progress streaming.
        """
        self.is_task_running = True
        self.message_queue.put(("update_buttons", "", ""))
        
        # Add status message
        self.message_queue.put(("Agent", "🤖 Initializing task...", "agent"))
        
        # CHANGE: Stream progress updates to GUI
        self.message_queue.put(("Agent", f"📋 Objective: {query}", "agent"))
        
        # Load and apply rules to the query
        rules = self.settings.get("rules", [])
        enhanced_query = query
        
        if rules:
            # Filter out empty rules
            valid_rules = [rule.strip() for rule in rules if rule.strip()]
            if valid_rules:
                self.message_queue.put(("Agent", f"📋 Applying {len(valid_rules)} rules to guide task execution...", "agent"))
                
                # Format rules as clear instructions
                rules_text = "IMPORTANT: Follow these rules while completing the task:\n"
                for i, rule in enumerate(valid_rules, 1):
                    rules_text += f"{i}. {rule}\n"
                rules_text += f"\nNow complete this task: {query}"
                enhanced_query = rules_text
        
        # Check for cancellation
        if asyncio.current_task().cancelled():
            raise asyncio.CancelledError()
        
        self.message_queue.put(("Agent", "🔍 Analyzing request and planning actions...", "agent"))
        
        # Execute agent query with progress streaming
        try:
            # Run the agent invoke in executor to avoid blocking
            loop = asyncio.get_event_loop()
            with ThreadPoolExecutor() as executor:
                self.message_queue.put(("Agent", "⚡ Executing agent task...", "agent"))
                
                # Check for cancellation before long operation
                if asyncio.current_task().cancelled():
                    raise asyncio.CancelledError()
                
                # Execute the actual agent task with output capture and rules
                agent_result = await loop.run_in_executor(executor, self.capture_agent_output, enhanced_query)
                
                # Check for cancellation after operation
                if asyncio.current_task().cancelled():
                    raise asyncio.CancelledError()
                
                # Display final result
                self.message_queue.put(("Agent", "✅ Task completed successfully!", "agent"))
                if hasattr(agent_result, 'content') and agent_result.content:
                    self.message_queue.put(("Agent", f"📄 Final Result: {agent_result.content}", "agent"))
                
        except asyncio.CancelledError:
            # Re-raise cancellation to be handled by caller
            raise
        except Exception as e:
            self.message_queue.put(("System", f"❌ Error executing task: {str(e)}", "error"))
        finally:
            self.is_task_running = False
            self.message_queue.put(("update_buttons", "", ""))
    
    def start_task(self):
        """Handle start task button click."""
        # Check if command is empty
        if not self.command_entry.get().strip() or self.command_entry.get() == "Message Luma":
            messagebox.showwarning("Warning", "Please enter a command first.")
            return
        
        # Check if API key is available
        if not self.current_api_key:
            self.add_message("System", "⚠️ API key is required to run tasks. Please enter your API key.", "error")
            self.show_settings()
            return
        
        # Initialize agent if not already initialized
        if not self.agent and DEPENDENCIES_AVAILABLE:
            try:
                self.initialize_agent()
            except Exception as e:
                self.add_message("System", f"❌ Failed to initialize agent: {str(e)}", "error")
                return
        
        # Check if agent is still not available after initialization attempt
        if not self.agent:
            self.add_message("System", "❌ Agent initialization failed. Please check your API key and try again.", "error")
            return
            
        # Proceed with task execution
        self.send_command()
    
    def stop_task(self):
        """CHANGE: Implement functional stop task with proper cancellation."""
        if self.is_task_running and self.task_thread:
            self.add_message("System", "🛑 Stopping task... Please wait...", "system")
            
            # Cancel the asyncio task if it exists
            if self.current_task and not self.current_task.done():
                self.current_task.cancel()
            
            # Reset state
            self.is_task_running = False
            self.update_button_states()
    
    def update_button_states(self):
        """Update button states based on task execution status."""
        if self.is_task_running:
            self.start_button.configure(state="disabled")
            self.stop_button.configure(state="normal")
        else:
            self.start_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
    
    def update_task_history_display(self):
        """Update the task history display in sidebar with proper styling."""
        # Clear existing history items
        for widget in self.task_history_frame.winfo_children():
            widget.destroy()
        
        # Reset selected task widget
        self.selected_task_widget = None
        
        # Add recent tasks (last 8 to fit better)
        recent_tasks = self.task_history[-8:] if len(self.task_history) > 8 else self.task_history
        
        for i, task in enumerate(reversed(recent_tasks)):
            task_text = task[:40] + "..." if len(task) > 40 else task
            
            task_label = tk.Label(
                self.task_history_frame,
                text=task_text,
                font=("Arial", 11),
                bg=self.colors['sidebar_bg'],
                fg=self.colors['fg'],
                anchor='w',
                cursor='hand2',
                padx=10,
                pady=8
            )
            task_label.pack(fill="x", pady=2)
            
            # Bind click event
            task_label.bind("<Button-1>", lambda e, t=task, widget=task_label: self.on_history_select(t, widget))
            
            # Hover effects
            def on_enter(event, widget=task_label):
                if widget != self.selected_task_widget:
                    widget.configure(bg=self.colors['button_bg'])
            
            def on_leave(event, widget=task_label):
                if widget != self.selected_task_widget:
                    widget.configure(bg=self.colors['sidebar_bg'])
            
            task_label.bind("<Enter>", on_enter)
            task_label.bind("<Leave>", on_leave)
        
        # Update scroll region
        self.task_history_frame.update_idletasks()
        self.task_canvas.configure(scrollregion=self.task_canvas.bbox("all"))
    
    def on_history_select(self, task, widget):
        """Handle task history item selection with visual feedback."""
        # Unhighlight previous selection
        if self.selected_task_widget:
            self.selected_task_widget.configure(bg=self.colors['sidebar_bg'])
        
        # Highlight current selection
        widget.configure(bg=self.colors['selected_bg'])
        self.selected_task_widget = widget
        
        # Load the task
        self.load_previous_task(task)
    
    def load_previous_task(self, task: str):
        """Load a previous task into the input field."""
        self.command_entry.delete(0, tk.END)
        self.command_entry.insert(0, task)
        self.command_entry.config(fg=self.colors['fg'])
    
    def process_message_queue(self):
        """
        Process messages from background threads safely in main thread.
        """
        try:
            while True:
                sender, message, msg_type = self.message_queue.get_nowait()
                
                if sender == "update_buttons":
                    self.update_button_states()
                else:
                    self.add_message(sender, message, msg_type)
                    
        except queue.Empty:
            pass
        
        # Schedule next check
        self.root.after(100, self.process_message_queue)
    
    def load_settings(self) -> Dict[str, Any]:
        """Load application settings from file."""
        settings_file = "agent_settings.json"
        default_settings = {
            "rules": [
                "Be helpful and accurate",
                "Provide clear explanations",
                "Ask for clarification when needed"
            ],
            "theme": "dark"
        }
        
        try:
            if os.path.exists(settings_file):
                with open(settings_file, 'r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading settings: {e}")
        
        return default_settings
    
    def save_settings_to_file(self):
        """Save current settings to file."""
        try:
            with open("agent_settings.json", 'w') as f:
                json.dump(self.settings, f, indent=2)
        except Exception as e:
            print(f"Error saving settings: {e}")
    
    def load_chat_history(self):
        """Load chat history from file."""
        history_file = "chat_history.json"
        try:
            if os.path.exists(history_file):
                with open(history_file, 'r') as f:
                    self.chat_history = json.load(f)
                
                # Display recent chat history (last 20 messages)
                recent_history = self.chat_history[-20:] if len(self.chat_history) > 20 else self.chat_history
                
                for msg in recent_history:
                    self.chat_display.configure(state=tk.NORMAL)
                    self.chat_display.insert(tk.END, f"[{msg['timestamp']}] {msg['sender']}: {msg['message']}\n\n")
                    self.chat_display.configure(state=tk.DISABLED)
                
                if recent_history:
                    self.chat_display.see(tk.END)
                    
        except Exception as e:
            print(f"Error loading chat history: {e}")
    
    def save_chat_history(self):
        """Save chat history to file."""
        try:
            with open("chat_history.json", 'w') as f:
                json.dump(self.chat_history, f, indent=2)
        except Exception as e:
            print(f"Error saving chat history: {e}")
    
    def load_task_history(self):
        """Load task history from file."""
        history_file = "task_history.json"
        try:
            if os.path.exists(history_file):
                with open(history_file, 'r') as f:
                    self.task_history = json.load(f)
                self.update_task_history_display()
        except Exception as e:
            print(f"Error loading task history: {e}")
    
    def save_task_history(self):
        """Save task history to file."""
        try:
            with open("task_history.json", 'w') as f:
                json.dump(self.task_history, f, indent=2)
        except Exception as e:
            print(f"Error saving task history: {e}")
    
    def open_machovotrish_link(self, event=None):
        """Open Machovotrish website in default browser."""
        try:
            webbrowser.open_new_tab("https://www.machovotrish.com")
        except Exception as e:
            # Fallback: show message box with URL if browser opening fails
            messagebox.showinfo("Website Link", "Visit: https://www.machovotrish.com")
            print(f"Error opening browser: {e}")
    
    def on_closing(self):
        """Handle application closing with cleanup."""
        try:
            # Save all data before closing
            self.save_chat_history()
            self.save_task_history()
            self.save_settings_to_file()
            
            # Stop any running tasks
            if self.is_task_running:
                self.is_task_running = False
            
            # Close application
            self.root.destroy()
            
        except Exception as e:
            print(f"Error during cleanup: {e}")
            self.root.destroy()
    
    def run(self):
        """Start the GUI application main loop."""
        self.root.mainloop()


def main():
    """
    Main entry point - launches GUI application instead of terminal interface.
    """
    try:
        # Check if dependencies are available
        if not DEPENDENCIES_AVAILABLE:
            print("Warning: Some dependencies are missing. GUI will start but agent functionality may be limited.")
            print("Please install: pip install langchain-google-genai windows-use python-dotenv")
        
        # Create and run GUI application
        app = WindowsAgentGUI()
        app.run()
        
    except Exception as e:
        print(f"Failed to start GUI application: {e}")
        # Fallback to original terminal interface if GUI fails
        if DEPENDENCIES_AVAILABLE:
            print("Falling back to terminal interface...")
            
            load_dotenv()
            llm = ChatGoogleGenerativeAI(model='gemini-2.0-flash')
            agent = Agent(llm=llm, browser='chrome', use_vision=True)
            
            query = input("Enter your query: ")
            agent_result = agent.invoke(query=query)
            print(agent_result.content)


if __name__ == "__main__":
    main()