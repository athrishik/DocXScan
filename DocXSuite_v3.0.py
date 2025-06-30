"""
Document Tools Suite v3.0 - Ultra Modern Launcher
Copyright 2025 Hrishik Kunduru. All rights reserved.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import platform
import ctypes
import subprocess
from datetime import datetime
import json
import os
import sys
from pathlib import Path


# DPI
if platform.system() == "Windows":
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
    except:
        pass

class ModernColors:
    """Premium dark theme color palette"""
    BG_PRIMARY = "#0A0E1A"
    BG_SECONDARY = "#0F1419"
    
    GLASS_BG = "#1A1D29"  # Changed from gray to dark blue-purple
    GLASS_BORDER = "#2A2D3A"  # Matching border
    GLASS_HOVER = "#1F2235"  # Hover state
    
    CARD_BG = "#161925"  # Dark blue instead of gray
    CARD_BORDER = "#252837"  # Subtle blue border
    CARD_HOVER = "#1C1F2E"  # Blue hover
    
    PRIMARY = "#2563EB"
    PRIMARY_HOVER = "#1D4ED8"
    SUCCESS = "#10B981"
    WARNING = "#F59E0B"
    ERROR = "#EF4444"
    
    TEXT_PRIMARY = "#FFFFFF"
    TEXT_SECONDARY = "#E5E7EB"
    TEXT_TERTIARY = "#9CA3AF"
    TEXT_MUTED = "#6B7280"
    
    FOCUS_RING = "#2563EB"

class BlurredBackground(tk.Canvas):
    
    def __init__(self, parent, width, height):
        super().__init__(parent, width=width, height=height, highlightthickness=0, bd=0)
        self.configure(bg=ModernColors.BG_PRIMARY)
        self.width = width
        self.height = height
        self.create_gradient_background()
    
    def create_gradient_background(self):
        """Create gradient background"""
        for i in range(self.height):
            ratio = i / self.height
            color_val = int(10 + (25 - 10) * (1 - ratio))
            color = f"#{color_val:02x}{color_val + 2:02x}{color_val + 8:02x}"
            self.create_line(0, i, self.width, i, fill=color, width=1)
        
        # Add decorative circles
        self.create_blur_circle(150, 100, 60)
        self.create_blur_circle(self.width - 120, 200, 40)
        self.create_blur_circle(300, self.height - 150, 80)

    def create_blur_circle(self, x, y, radius):
        """Create decorative circle"""
        colors = [ModernColors.PRIMARY, ModernColors.SUCCESS]
        for i, color in enumerate(colors[:2]):
            r = radius - i * 15
            if r > 0:
                self.create_oval(x - r, y - r, x + r, y + r,
                               fill=color, outline="", stipple="gray12")

class GlassMorphCard(tk.Frame):

    def __init__(self, parent, title, subtitle, icon, primary_color, command, shortcut=""):
        super().__init__(parent)
        
        self.title = title
        self.subtitle = subtitle
        self.icon = icon
        self.primary_color = primary_color
        self.command = command
        self.shortcut = shortcut
        self.font_family = self.get_system_font()
        
        self.setup_card()
        self.create_content()
        self.setup_interactions()
    
    def get_system_font(self):
        system = platform.system()
        fonts = {
            "Darwin": "SF Pro Display",
            "Windows": "Segoe UI",
            "Linux": "Ubuntu"
        }
        return fonts.get(system, "Arial")
    
    def check_programs_available(self):
        """Check if both programs are available"""
        scan_exists = os.path.exists("docxscan2_0.py")
        replace_exists = os.path.exists("docxreplace2_0.py")
        return scan_exists and replace_exists
    
    def get_status_color_and_text(self):
        """Get status based on program availability"""
        if self.check_programs_available():
            return ModernColors.SUCCESS, "All systems operational"
        else:
            return ModernColors.WARNING, "Some programs not detected"
    
    def setup_card(self):
        """Setup card styling"""
        self.configure(
            bg=ModernColors.CARD_BG,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=ModernColors.CARD_BORDER
        )
        self.configure(width=420, height=300)
        self.pack_propagate(False)
    
    def create_content(self):
        """Create card content"""
        container = tk.Frame(self, bg=ModernColors.CARD_BG)
        container.pack(fill=tk.BOTH, expand=True, padx=28, pady=24)
        
        # Header with icon - reduced top padding
        header_frame = tk.Frame(container, bg=ModernColors.CARD_BG)
        header_frame.pack(fill=tk.X, pady=(0, 16))
        
        # Icon
        icon_container = tk.Frame(header_frame, bg=self.primary_color, width=60, height=60)
        icon_container.pack_propagate(False)
        icon_container.pack(side=tk.LEFT)
        
        icon_label = tk.Label(icon_container, text=self.icon,
                             font=(self.font_family, 26, "normal"),
                             bg=self.primary_color, fg="white")
        icon_label.pack(expand=True)
        
        # Status dot
        status_dot = tk.Label(header_frame, text="●",
                             font=("Arial", 8),
                             fg=ModernColors.SUCCESS,
                             bg=ModernColors.CARD_BG)
        status_dot.pack(side=tk.RIGHT)
        
        # Title
        title_label = tk.Label(container, text=self.title,
                              font=(self.font_family, 22, "bold"),
                              bg=ModernColors.CARD_BG,
                              fg=ModernColors.TEXT_PRIMARY,
                              anchor="w")
        title_label.pack(fill=tk.X, pady=(0, 6))
        
        # Subtitle
        subtitle_label = tk.Label(container, text=self.subtitle,
                                 font=(self.font_family, 12, "normal"),
                                 bg=ModernColors.CARD_BG,
                                 fg=ModernColors.TEXT_SECONDARY,
                                 wraplength=360,
                                 justify="left",
                                 anchor="w")
        subtitle_label.pack(fill=tk.X, pady=(0, 20))
        
        # Action section
        action_frame = tk.Frame(container, bg=ModernColors.CARD_BG)
        action_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        # launch button
        self.action_btn = tk.Button(action_frame,
                                   text=f"Launch {self.title}",
                                   font=(self.font_family, 13, "bold"),
                                   bg=self.primary_color,
                                   fg="white",
                                   activebackground=self.primary_color,
                                   activeforeground="white",
                                   bd=0,
                                   relief="flat",
                                   padx=40,
                                   pady=18,
                                   cursor="hand2",
                                   command=self.command)
        self.action_btn.pack(side=tk.LEFT)
        
        # Shortcut badge with improved styling
        if self.shortcut:
            shortcut_frame = tk.Frame(action_frame, bg=ModernColors.GLASS_BG,
                                     highlightthickness=1,
                                     highlightbackground=ModernColors.GLASS_BORDER)
            shortcut_frame.pack(side=tk.RIGHT, padx=(16, 0))
            
            shortcut_label = tk.Label(shortcut_frame, text=self.shortcut,
                                     font=(self.font_family, 11, "bold"),
                                     bg=ModernColors.GLASS_BG,
                                     fg=ModernColors.TEXT_TERTIARY,
                                     padx=12, pady=8)
            shortcut_label.pack()
        
        # Store widgets for hover effects
        self.hover_widgets = [self, container, header_frame, title_label, subtitle_label]
    
    def setup_interactions(self):
        """Setup hover interactions - only for card areas, not button"""
        for widget in self.hover_widgets:
            widget.bind("<Enter>", self.on_enter)
            widget.bind("<Leave>", self.on_leave)
        
        self.action_btn.bind("<Enter>", self.on_button_enter)
        self.action_btn.bind("<Leave>", self.on_button_leave)
    
    def on_enter(self, event):
        """Hover effect - no background highlighting"""
        self.configure(highlightbackground=ModernColors.FOCUS_RING)
    
    def on_leave(self, event):
        """Reset hover"""
        self.configure(highlightbackground=ModernColors.CARD_BORDER)
    
    def on_button_enter(self, event):
        """Button hover"""
        self.action_btn.configure(bg=ModernColors.PRIMARY_HOVER)
    
    def on_button_leave(self, event):
        """Button reset"""
        self.action_btn.configure(bg=self.primary_color)

class ModernStatsPanel(tk.Frame):
    """Statistics panel with glass styling"""
    
    def __init__(self, parent, usage_data, font_family):
        super().__init__(parent, bg=ModernColors.BG_PRIMARY)
        self.usage_data = usage_data
        self.font_family = font_family
        self.create_stats_panel()
    
    def create_stats_panel(self):
        """Create stats panel"""
        stats_container = tk.Frame(self,
                                  bg=ModernColors.GLASS_BG,
                                  relief="flat",
                                  bd=0,
                                  highlightthickness=1,
                                  highlightbackground=ModernColors.GLASS_BORDER)
        stats_container.pack(fill=tk.X, padx=20, pady=10)
        
        # Header
        header_frame = tk.Frame(stats_container, bg=ModernColors.GLASS_BG)
        header_frame.pack(fill=tk.X, padx=24, pady=(18, 12))
        
        tk.Label(header_frame, text="Usage Analytics",
                font=(self.font_family, 14, "bold"),
                bg=ModernColors.GLASS_BG,
                fg=ModernColors.TEXT_PRIMARY).pack(side=tk.LEFT)
        
        # Live indicator
        live_frame = tk.Frame(header_frame, bg=ModernColors.GLASS_BG)
        live_frame.pack(side=tk.RIGHT)
        
        tk.Label(live_frame, text="● LIVE",
                font=(self.font_family, 9, "bold"),
                fg=ModernColors.SUCCESS,
                bg=ModernColors.GLASS_BG).pack()
        
        # Stats grid
        stats_grid = tk.Frame(stats_container, bg=ModernColors.GLASS_BG)
        stats_grid.pack(fill=tk.X, padx=24, pady=(0, 18))
        
        for i in range(4):
            stats_grid.grid_columnconfigure(i, weight=1)
        
        # Create stats
        stats = [
            ("Scans", self.usage_data.get("scan_count", 0), ModernColors.PRIMARY),
            ("Replacements", self.usage_data.get("replace_count", 0), ModernColors.SUCCESS),
            ("Sessions", self.usage_data.get("total_sessions", 0), ModernColors.WARNING),
            ("Last Used", self.format_last_used(), ModernColors.TEXT_TERTIARY)
        ]
        
        for i, (label, value, color) in enumerate(stats):
            self.create_stat_item(stats_grid, label, str(value), color, i)
    
    def format_last_used(self):
        """Format last used date"""
        last_used = self.usage_data.get("last_used", "Never")
        if last_used == "Never":
            return "Never"
        try:
            date_obj = datetime.fromisoformat(last_used)
            return date_obj.strftime("%b %d")
        except:
            return "Never"
    
    def create_stat_item(self, parent, label, value, color, column):
        """Create stat item"""
        stat_frame = tk.Frame(parent, bg=ModernColors.GLASS_BG)
        stat_frame.grid(row=0, column=column, padx=12, pady=8, sticky="ew")
        
        tk.Label(stat_frame, text=value,
                font=(self.font_family, 24, "bold"),
                bg=ModernColors.GLASS_BG,
                fg=color).pack()
        
        tk.Label(stat_frame, text=label,
                font=(self.font_family, 10, "normal"),
                bg=ModernColors.GLASS_BG,
                fg=ModernColors.TEXT_MUTED).pack(pady=(2, 0))

class UltraModernLauncher:
    """Ultra-modern launcher application"""
    
    def __init__(self):
        self.root = None
        self.font_family = self.get_system_font()
        self.usage_data = self.load_usage_data()
        
    def get_system_font(self):
        system = platform.system()
        fonts = {
            "Darwin": "SF Pro Display",
            "Windows": "Segoe UI",
            "Linux": "Ubuntu"
        }
        return fonts.get(system, "Arial")
    
    def load_usage_data(self):
        usage_file = Path("usage_stats.json")
        default_data = {
            "scan_count": 0,
            "replace_count": 0,
            "last_used": "Never",
            "total_sessions": 0
        }
        
        if usage_file.exists():
            try:
                with open(usage_file, 'r') as f:
                    data = json.load(f)
                    return {**default_data, **data}
            except:
                pass
        
        return default_data
    
    def save_usage_data(self):
        try:
            with open("usage_stats.json", 'w') as f:
                json.dump(self.usage_data, f, indent=2)
        except:
            pass
    
    def create_window(self):
        """Create modern window"""
        if self.root:
            try:
                self.root.destroy()
            except:
                pass
        
        self.root = tk.Tk()
        self.root.title("Document Tools Suite")
        self.root.geometry("1200x900")
        self.root.configure(bg=ModernColors.BG_PRIMARY)
        self.root.resizable(True, True)
        self.root.minsize(1000, 800)
        
        if platform.system() == "Windows":
            try:
                self.root.wm_attributes("-alpha", 0.98)
            except:
                pass
        
        self.create_interface()
        self.setup_shortcuts()
        self.center_window()
    
    def create_interface(self):
        """Create interface"""
        bg_canvas = BlurredBackground(self.root, 1200, 900)
        bg_canvas.pack(fill=tk.BOTH, expand=True)
        
        main_overlay = tk.Frame(bg_canvas, bg=ModernColors.BG_PRIMARY)
        main_overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
        
        self.create_header(main_overlay)
        self.create_tools_section(main_overlay)
        self.create_stats_section(main_overlay)
        self.create_footer(main_overlay)
    
    def create_header(self, parent):
        """Create header"""
        header_frame = tk.Frame(parent, bg=ModernColors.BG_PRIMARY, height=140)
        header_frame.pack(fill=tk.X, padx=60, pady=(40, 20))
        header_frame.pack_propagate(False)
        
        title_container = tk.Frame(header_frame, bg=ModernColors.BG_PRIMARY)
        title_container.pack(expand=True, fill=tk.BOTH)
        
        tk.Label(title_container,
                text="Document Tools Suite",
                font=(self.font_family, 32, "bold"),
                bg=ModernColors.BG_PRIMARY,
                fg=ModernColors.TEXT_PRIMARY).pack(anchor="w")
        
        tk.Label(title_container,
                text="Professional legal document processing platform",
                font=(self.font_family, 14, "normal"),
                bg=ModernColors.BG_PRIMARY,
                fg=ModernColors.TEXT_SECONDARY).pack(anchor="w", pady=(4, 0))
        
        # Version and status
        meta_frame = tk.Frame(title_container, bg=ModernColors.BG_PRIMARY)
        meta_frame.pack(anchor="w", pady=(16, 0), fill=tk.X)
        
        version_badge = tk.Frame(meta_frame,
                                bg=ModernColors.GLASS_BG,
                                highlightthickness=1,
                                highlightbackground=ModernColors.GLASS_BORDER)
        version_badge.pack(side=tk.LEFT)
        
        tk.Label(version_badge, text="v3.0",
                font=(self.font_family, 11, "bold"),
                bg=ModernColors.GLASS_BG,
                fg=ModernColors.PRIMARY,
                padx=12, pady=6).pack()
        
        status_frame = tk.Frame(meta_frame, bg=ModernColors.BG_PRIMARY)
        status_frame.pack(side=tk.LEFT, padx=(20, 0))
        
        # Dynamic status
        try:
            status_color, status_text = self.get_status_color_and_text()
        except AttributeError:
            status_color, status_text = ModernColors.SUCCESS, "All systems operational"
        
        tk.Label(status_frame, text=f"● {status_text}",
                font=(self.font_family, 11, "normal"),
                fg=status_color,
                bg=ModernColors.BG_PRIMARY).pack()
    
    def create_tools_section(self, parent):
        """Create tools section"""
        tools_frame = tk.Frame(parent, bg=ModernColors.BG_PRIMARY)
        tools_frame.pack(fill=tk.BOTH, expand=True, padx=60, pady=(20, 30))
        
        grid_container = tk.Frame(tools_frame, bg=ModernColors.BG_PRIMARY)
        grid_container.pack(fill=tk.X, pady=20)
        
        grid_container.grid_columnconfigure(0, weight=1)
        grid_container.grid_columnconfigure(1, weight=1)
        
        # DocXScan card
        scan_card = GlassMorphCard(
            grid_container,
            title="DocXScan",
            subtitle="Advanced document scanner with intelligent token detection, pattern matching, and comprehensive analysis capabilities",
            icon="🔍",
            primary_color=ModernColors.PRIMARY,
            command=self.launch_scan,
            shortcut="F1"
        )
        scan_card.grid(row=0, column=0, padx=(0, 15), pady=0, sticky="nsew")
        
        # DocXReplace card
        replace_card = GlassMorphCard(
            grid_container,
            title="DocXReplace",
            subtitle="Professional document token replacement with advanced regex support, batch processing, and intelligent pattern migration",
            icon="🔄",
            primary_color=ModernColors.SUCCESS,
            command=self.launch_replace,
            shortcut="F2"
        )
        replace_card.grid(row=0, column=1, padx=(15, 0), pady=0, sticky="nsew")
    
    def create_stats_section(self, parent):
        """Create stats section"""
        stats_frame = tk.Frame(parent, bg=ModernColors.BG_PRIMARY)
        stats_frame.pack(fill=tk.X, padx=60, pady=(0, 30))
        
        stats_panel = ModernStatsPanel(stats_frame, self.usage_data, self.font_family)
        stats_panel.pack(fill=tk.X)
    
    def create_footer(self, parent):
        """Create footer"""
        footer_frame = tk.Frame(parent, bg=ModernColors.BG_PRIMARY)
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=60, pady=(0, 30))
        
        tk.Label(footer_frame,
                text="© 2025 Hrishik Kunduru • Document Tools Suite • All Rights Reserved",
                font=(self.font_family, 10, "normal"),
                bg=ModernColors.BG_PRIMARY,
                fg=ModernColors.TEXT_MUTED).pack()
    
    def setup_shortcuts(self):
        """Setup shortcuts"""
        self.root.bind('<F1>', lambda e: self.launch_scan())
        self.root.bind('<F2>', lambda e: self.launch_replace())
        self.root.bind('<Escape>', lambda e: self.close_launcher())
        self.root.bind('<Control-r>', lambda e: self.reset_stats())
        self.root.focus_force()
    
    def reset_stats(self):
        """Reset statistics"""
        if messagebox.askyesno("Reset Statistics", "Reset all usage statistics to zero?"):
            self.usage_data = {
                "scan_count": 0,
                "replace_count": 0,
                "last_used": "Never",
                "total_sessions": 1
            }
            self.save_usage_data()
            self.create_window()
    
    def center_window(self):
        """Center window"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def close_launcher(self):
        """Close launcher"""
        try:
            self.save_usage_data()
            if self.root:
                self.root.destroy()
                self.root = None
        except:
            pass
    
    def run(self):
        """Run launcher"""
        try:
            self.usage_data["total_sessions"] += 1
            self.save_usage_data()
            
            self.create_window()
            self.root.protocol("WM_DELETE_WINDOW", self.close_launcher)
            self.root.mainloop()
        except Exception as e:
            print(f"Launcher error: {e}")
        finally:
            self.close_launcher()
    
    def launch_scan(self):
        """Launch DocXScan EXE and close launcher"""
        try:
            self.usage_data["scan_count"] += 1
            self.usage_data["last_used"] = datetime.now().isoformat()
            self.save_usage_data()
            
            # Launch the EXE from its folder
            subprocess.Popen(["DocXScan/DocXScan.exe"])
            
            print("DocXScan launched successfully - closing launcher")
            
            # Close the launcher
            self.close_launcher()
            
        except Exception as e:
            messagebox.showerror("Launch Error", f"Failed to launch DocXScan:\n{str(e)}")
    
    def launch_replace(self):
        """Launch DocXReplace EXE and close launcher"""
        try:
            self.usage_data["replace_count"] += 1
            self.usage_data["last_used"] = datetime.now().isoformat()
            self.save_usage_data()
            
            # Launch the EXE from its folder
            subprocess.Popen(["DocXReplace/DocXReplace.exe"])
            
            print("DocXReplace launched successfully - closing launcher")
            
            # Close the launcher
            self.close_launcher()
            
        except Exception as e:
            messagebox.showerror("Launch Error", f"Failed to launch DocXReplace:\n{str(e)}")

def main():
    """Main entry point"""
    try:
        print("Starting Document Tools Suite...")
        app = UltraModernLauncher()
        app.run()
    except Exception as e:
        error_msg = f"Critical startup error: {str(e)}\n\nTraceback:\n{traceback.format_exc()}"
        print(error_msg)
        # Try to show error dialog if possible
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Startup Error", error_msg)
        except:
            print("Could not display error dialog")
        input("Press Enter to exit...")  # Keep console open

if __name__ == "__main__":
    main()


