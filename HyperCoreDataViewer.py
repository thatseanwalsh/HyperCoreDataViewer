__author__ = "Sean Michael Walsh"
__credits__ = "The Scale Computing Solutions Architect Team"
__maintainer__ = "Sean Michael Walsh"
__email__ = "swalsh@scalecomputing.com"
__status__ = "Development"

"""
HyperCoreDataViewer.py: An application to use the Scale Computing 
                        HyperCore API to export cluster and virtual 
                        machine data for the purpose of viewing and 
                        exporting to a spreadsheet.
"""

import customtkinter as ctk
from tkinter import Toplevel, ttk, filedialog, VERTICAL, HORIZONTAL
from PIL import Image, ImageTk 
import base64
import http.client as http
import json
import ssl
import pandas as pd
import xml.etree.ElementTree as ET
from tkinter import font
import os
import sys
import platform

# Global dark mode
ctk.set_appearance_mode("dark")

class ClusterApp:
    def __init__(self, root):
        self.root = root
        root.focus_force()

        self.cluster_ip = ""
        self.username = ""
        self.password = ""
        self.processed_vms = {}
        self.processed_cluster = {}
        self.setup_gui()

    # GUI setup, logos, icons, and the like
    def setup_gui(self):
        def resource_path(relative_path):
            if getattr(sys, 'frozen', False):  # Compiled bundle check
                base_path = os.path.dirname(sys.executable)

                if platform.system() == "Windows":
                    base_path = sys._MEIPASS
                    return os.path.join(base_path, "assets", relative_path)  # Adjust for Windows
                elif platform.system() == "Darwin": 
                    return os.path.join(base_path, "../Resources", relative_path)  # Adjust for macOS
            else:
                return os.path.join(os.path.abspath("./assets"), relative_path) # Adjust for CLI
        
        self.root.title("SC//HyperCore Data Viewer")
        if platform.system() == "Windows":
            icon_path = resource_path("icon.ico")  # .ico for Windows
            self.root.iconbitmap(icon_path)
        elif platform.system() == "Darwin":
            icon_path = resource_path("icon.icns")  # .icns for macOS
            icon_image = Image.open(icon_path)
            icon_photo = ImageTk.PhotoImage(icon_image)
            self.root.iconphoto(True, icon_photo)
        icon_png_path = resource_path("icon.png")
        logo_path = resource_path("logo.png")
        icon_image = Image.open(icon_png_path)

        ctk.set_appearance_mode("dark")
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = 1024
        window_height = 768
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

        self.top_frame = ctk.CTkFrame(self.root)
        self.top_frame.pack(fill=ctk.X, padx=20, pady=(20, 5)) 

        logo_image = ctk.CTkImage(light_image=Image.open(logo_path), size=(220, 93)) 
        self.logo_label = ctk.CTkLabel(self.top_frame, image=logo_image, text="")
        self.logo_label.image = logo_image 
        self.logo_label.pack(side=ctk.LEFT, padx=10, pady=10)

        self.button_frame_top = ctk.CTkFrame(self.root)
        self.button_frame_top.pack(pady=(5,0))

        # View buttons
        self.view_button1 = ctk.CTkButton(self.button_frame_top, text="Cluster", command=self.switch_view_cluster, font=("Martel Sans", 14), fg_color="#e3004b", hover_color="#e67b34")
        self.view_button1.pack(side=ctk.LEFT, padx=10, pady=10)

        self.view_button2 = ctk.CTkButton(self.button_frame_top, text="Node", command=self.switch_view_node, font=("Martel Sans", 14), fg_color="#e3004b", hover_color="#e67b34")
        self.view_button2.pack(side=ctk.LEFT, padx=10, pady=10)

        self.view_button3 = ctk.CTkButton(self.button_frame_top, text="Virtual Machines", command=self.switch_view_vm, font=("Martel Sans", 14), fg_color="#e3004b", hover_color="#e67b34")
        self.view_button3.pack(side=ctk.LEFT, padx=10, pady=10)

        # Instructions box
        self.instructions = ctk.CTkTextbox(self.top_frame, height=150, width=500, font=("Martel Sans", 12), padx=10)
        self.instructions.insert(ctk.END, "Instructions:\n1. Create a read-only user in your SC//HyperCore user interface.\n"
                                        "2. Click 'Settings' to enter cluster credentials.\n"
                                        "3. Click 'Fetch Data' to load the cluster's information.\n"
                                        "4. Click 'Export' to export the collected data to a spreadsheet.\n"
                                        "This tool is not endorsed or supported by Scale Computing. Use at your own risk.")
        self.instructions.configure(state="disabled") 
        self.instructions.pack(side=ctk.RIGHT, padx=10, pady=10, fill=ctk.X, expand=True)

        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(fill=ctk.BOTH, expand=True, padx=20, pady=(10,10))

        self.cluster_frame = ctk.CTkFrame(self.main_frame)
        self.cluster_frame.pack(fill=ctk.BOTH, expand=True)
        self.cluster_tree = ttk.Treeview(self.cluster_frame, show="headings")

        self.node_frame = ctk.CTkFrame(self.main_frame)
        self.node_frame.pack(fill=ctk.BOTH, expand=True)
        self.node_tree = ttk.Treeview(self.node_frame, show="headings")

        self.vm_frame = ctk.CTkFrame(self.main_frame)
        self.vm_frame.pack(fill=ctk.BOTH, expand=True)
        self.vm_tree = ttk.Treeview(self.vm_frame, show="headings")

        # Create scrollbars
        self.vm_tree_scroll_x = ctk.CTkScrollbar(self.vm_frame, orientation="horizontal", command=self.vm_tree.xview)
        self.vm_tree_scroll_x.pack(side=ctk.BOTTOM, fill=ctk.X)
        self.vm_tree_scroll_y = ctk.CTkScrollbar(self.vm_frame, orientation="vertical", command=self.vm_tree.yview)
        self.vm_tree_scroll_y.pack(side=ctk.RIGHT, fill=ctk.Y)
        self.cluster_tree_scroll_y = ctk.CTkScrollbar(self.cluster_frame, orientation="vertical", command=self.cluster_tree.yview)
        self.cluster_tree_scroll_y.pack(side=ctk.RIGHT, fill=ctk.Y)
        self.node_tree_scroll_y = ctk.CTkScrollbar(self.node_frame, orientation="vertical", command=self.node_tree.yview)
        self.node_tree_scroll_y.pack(side=ctk.RIGHT, fill=ctk.Y)
        
        # Create treeview
        self.vm_tree = ttk.Treeview(
            self.vm_frame, show="headings", 
            xscrollcommand=self.vm_tree_scroll_x.set, 
            yscrollcommand=self.vm_tree_scroll_y.set
        )
        self.cluster_tree = ttk.Treeview(self.cluster_frame, show="headings", yscrollcommand=self.cluster_tree_scroll_y.set)
        self.node_tree = ttk.Treeview(self.node_frame, show="headings", yscrollcommand=self.node_tree_scroll_y.set)

        # Pack the treeview 
        self.cluster_tree.pack(fill=ctk.BOTH, expand=True, padx=(0, 10), pady=(0, 10))
        self.node_tree.pack(fill=ctk.BOTH, expand=True, padx=(0, 10), pady=(0, 10))
        self.vm_tree.pack(fill=ctk.BOTH, expand=True, padx=(0, 10), pady=(0, 10))

        # Configure treeview scrollbars
        self.vm_tree_scroll_x.configure(command=self.vm_tree.xview)
        self.vm_tree_scroll_y.configure(command=self.vm_tree.yview)
        self.cluster_tree_scroll_y.configure(command=self.cluster_tree.yview)
        self.node_tree_scroll_y.configure(command=self.node_tree.yview)

        # Treeview styling
        style = ttk.Style()
        style.theme_use("default")  
        style.configure("Treeview.Heading",
                background="#3e3e3e",
                foreground="white",
                font=("Martel Sans", 14),
                borderwidth=0,
                anchor="w",
                padding=(5, 5))
        
        style.map("Treeview.Heading",
                background=[("selected", "#e3004b")],
                foreground=[("selected", "#ffffff")])
        
        style.configure("Treeview",
                background="#2e2e2e",
                foreground="#ffffff",
                rowheight=30,
                fieldbackground="#2e2e2e",
                font=("Martel Sans", 14),
                borderwidth=0,
                padding=(5, 5))

        style.map("Treeview",
                background=[("selected", "#e67b34")],
                foreground=[("selected", "#ffffff")])

        self.cluster_tree.pack(fill=ctk.BOTH, expand=True, padx=10, pady=10)
        self.cluster_tree.tag_configure('oddrow', background='#2e2e2e')
        self.cluster_tree.tag_configure('evenrow', background='#1e1e1e')
        self.node_tree.pack(fill=ctk.BOTH, expand=True, padx=10, pady=10)
        self.node_tree.tag_configure('oddrow', background='#2e2e2e')
        self.node_tree.tag_configure('evenrow', background='#1e1e1e')
        self.vm_tree.pack(fill=ctk.BOTH, expand=True, padx=10, pady=10)
        self.vm_tree.tag_configure('oddrow', background='#2e2e2e')
        self.vm_tree.tag_configure('evenrow', background='#1e1e1e')

        self.cluster_frame.pack_forget()
        self.vm_frame.pack_forget()

        self.button_frame = ctk.CTkFrame(self.root)
        self.button_frame.pack(pady=(0,10))

        self.settings_button = ctk.CTkButton(self.button_frame, text="Settings", command=self.open_settings, font=("Martel Sans", 14), fg_color="#e3004b", hover_color="#e67b34")
        self.settings_button.pack(side=ctk.LEFT, padx=10, pady=10)

        self.export_button = ctk.CTkButton(self.button_frame, text="Export", command=self.export, font=("Martel Sans", 14), fg_color="#e3004b", hover_color="#e67b34")
        self.export_button.pack(side=ctk.LEFT, padx=10, pady=10)

    def switch_view_cluster(self):
        self.fetch_data(view_type="cluster")
        self.cluster_frame.pack(fill=ctk.BOTH, expand=True)
        self.vm_frame.pack_forget() 
        self.node_frame.pack_forget() 

    def switch_view_node(self):
        self.fetch_data(view_type="node")
        self.node_frame.pack(fill=ctk.BOTH, expand=True)
        self.vm_frame.pack_forget()
        self.cluster_frame.pack_forget()

    def switch_view_vm(self):
        self.fetch_data(view_type="vm")
        self.vm_frame.pack(fill=ctk.BOTH, expand=True)
        self.cluster_frame.pack_forget()
        self.node_frame.pack_forget()

    # Settings modal
    def open_settings(self, view_type="cluster"):
        self.current_view = view_type 
        settings_window = Toplevel(self.root)
        settings_window.title("Settings")
        settings_window.transient(self.root)
        settings_window.focus_set()
        settings_window.grab_set()
        ctk.set_appearance_mode("dark")

        window_width = 300
        window_height = 230
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width  // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        settings_window.geometry(f"{window_width}x{window_height}+{x}+{y}")

        frame = ctk.CTkFrame(settings_window)
        frame.pack(pady=20, padx=20, fill='both', expand=True)

        ctk.CTkLabel(frame, text="Node IP:", font=("Martel Sans", 14)).grid(row=0, column=0, padx=10, pady=(15, 5), sticky='w')
        node_ip_entry = ctk.CTkEntry(frame, font=("Martel Sans", 14))
        node_ip_entry.grid(row=0, column=1, padx=10, pady=(15, 5), sticky='ew')
        node_ip_entry.insert(0, self.cluster_ip)
        node_ip_entry.focus()

        ctk.CTkLabel(frame, text="Username:", font=("Martel Sans", 14)).grid(row=1, column=0, padx=10, pady=5, sticky='w')
        username_entry = ctk.CTkEntry(frame, font=("Martel Sans", 14))
        username_entry.grid(row=1, column=1, padx=10, pady=5, sticky='ew')
        username_entry.insert(0, self.username)

        ctk.CTkLabel(frame, text="Password:", font=("Martel Sans", 14)).grid(row=2, column=0, padx=10, pady=5, sticky='w')
        password_entry = ctk.CTkEntry(frame, font=("Martel Sans", 14), show="*")
        password_entry.grid(row=2, column=1, padx=10, pady=5, sticky='ew')
        password_entry.insert(0, self.password)

        def save_settings():
            self.cluster_ip = node_ip_entry.get()
            self.username = username_entry.get()
            self.password = password_entry.get()
            settings_window.destroy()
            self.fetch_data(self.current_view)

        save_button = ctk.CTkButton(frame, text="Fetch Data", command=save_settings, font=("Martel Sans", 14), fg_color="#e3004b", hover_color="#e67b34")
        save_button.grid(row=3, column=0, columnspan=2, pady=10)
        settings_window.bind("<Return>", lambda event: save_settings())
        
        frame.columnconfigure(1, weight=1)

    def update_cluster_columns(self, columns):
        for col in self.cluster_tree["columns"]:
            self.cluster_tree.heading(col, text="")
        
        self.cluster_tree["columns"] = columns
        for col in columns:
            self.cluster_tree.heading(col, text=col, anchor="w")
            if col in ["Tag"]:
                self.cluster_tree.column(col, anchor="w", width=200, stretch=False)
            else:
                self.cluster_tree.column(col, anchor="w", width=750, stretch=True)
        
        for item in self.cluster_tree.get_children():
            self.cluster_tree.delete(item)

    def update_node_columns(self, columns):
        for col in self.node_tree["columns"]:
            self.node_tree.heading(col, text="")
        
        self.node_tree["columns"] = columns
        for col in columns:
            self.node_tree.heading(col, text=col, anchor="w")
            if col in ["TO BE DEFINED"]:
                self.node_tree.column(col, anchor="w", width=200, stretch=False)
            else:
                self.node_tree.column(col, anchor="w", width=750, stretch=True)
        
        for item in self.node_tree.get_children():
            self.node_tree.delete(item)     

    def update_vm_columns(self, columns):
        for col in self.vm_tree["columns"]:
            self.vm_tree.heading(col, text="", anchor="w")
        self.processed_vms.clear()  
        
        self.vm_tree["columns"] = columns
        tree_font = font.nametofont("TkDefaultFont")
        for col in columns:
            self.vm_tree.heading(col, text=col, command=lambda _col=col: self.sort_vm_tree(_col), anchor="w")
            text_width = tree_font.measure(col) + 20
            if col in ["Capacity (GiB)", "Allocation (GiB)"]:
                self.vm_tree.column(col, anchor="e", stretch=False, width=text_width)
            elif col in ["vCPUs", "Memory (GiB)"]:
                self.vm_tree.column(col, anchor="center", stretch=False, width=text_width)
            else:
                self.vm_tree.column(col, anchor="w", stretch=False, width=text_width)
        
        for item in self.vm_tree.get_children():
            self.vm_tree.delete(item)

    def sort_vm_tree(self, col, reverse=False):
        children = self.vm_tree.get_children("")
        data = []
        total_item = None

        for item in children:
            row_values = self.vm_tree.item(item, "values")
            if "TOTAL" in row_values:  # Identify the TOTAL row
                total_item = item
            else:
                data.append((self.vm_tree.set(item, col), item))

        # Sort data excluding total row
        data.sort(key=lambda x: (float(x[0]) if x[0].replace('.', '', 1).isdigit() else x[0].lower()), reverse=reverse)

        # Move sorted rows back into the tree
        for index, (val, item) in enumerate(data):
            self.vm_tree.move(item, "", index)

        # Ensure total row sits at the bottom
        if total_item:
            self.vm_tree.move(total_item, "", len(data))

        # Reapply alternating row colors
        self.alternate_row_colors()

        # Reapply total row styling
        if total_item:
            self.vm_tree.item(total_item, tags=('total',))

        self.vm_tree.tag_configure('total', background='#e3004b', foreground='#ffffff')

        # Toggle sorting direction
        self.vm_tree.heading(col, command=lambda: self.sort_vm_tree(col, not reverse))

    def alternate_row_colors(self):
        children = self.vm_tree.get_children()
        
        for index, item in enumerate(self.cluster_tree.get_children()):
            tag = 'evenrow' if index % 2 == 0 else 'oddrow'
            self.cluster_tree.item(item, tags=(tag,))

        for index, item in enumerate(self.node_tree.get_children()):
            tag = 'evenrow' if index % 2 == 0 else 'oddrow'
            self.node_tree.item(item, tags=(tag,))

        for index, item in enumerate(self.vm_tree.get_children()):
            tag = 'evenrow' if index % 2 == 0 else 'oddrow'
            self.vm_tree.item(item, tags=(tag,))
    
    # Error/message box 
    def show_message_box(self, error_message):
        message_box = Toplevel(self.root)
        message_box.title("Message")
        message_box.transient(self.root)
        message_box.grab_set()
        message_box.focus_set()
        width = 250
        height = 150
        x = (message_box.winfo_screenwidth() // 2) - (width // 2)
        y = (message_box.winfo_screenheight() // 2) - (height // 2)
        message_box.geometry(f"{width}x{height}+{x}+{y}")
        ctk.set_appearance_mode("dark")

        frame = ctk.CTkFrame(message_box)
        frame.pack(pady=20, padx=20, fill='both')

        label = ctk.CTkLabel(frame, text=error_message, font=("Martel Sans", 14), wraplength=200)
        label.pack(pady=5, padx=5)

        button = ctk.CTkButton(message_box, text="OK", command=message_box.destroy, font=("Martel Sans", 14), fg_color="#e3004b", hover_color="#e67b34")
        button.pack(pady=5)
        message_box.bind("<Return>", lambda event: message_box.destroy())

        message_box.protocol("WM_DELETE_WINDOW", message_box.destroy)

    # Data fetch
    def fetch_data(self, view_type):
        host = self.cluster_ip
        username = self.username
        password = self.password

        if not host or not username or not password:
            self.open_settings(view_type)
            return

        try:
            if view_type == "cluster":
                self.columns_cluster = ("Tag", "Value")
                self.update_cluster_columns(self.columns_cluster)
                self.vm_frame.pack_forget()
                self.cluster_frame.pack(fill=ctk.BOTH, expand=True)
                self.fetch_cluster_data(host, username, password)
            elif view_type == "vm":
                self.columns_vm = (
                "Name", "UUID", "Description", "OS", "Machine Type", "State", "vCPUs", "Memory (GiB)", 
                "Block Device", "Device Type", "Capacity (GiB)", "Allocation (GiB)", "Mount Points"
                )
                self.update_vm_columns(self.columns_vm)
                self.cluster_frame.pack_forget()  
                self.vm_frame.pack(fill=ctk.BOTH, expand=True)
                self.fetch_vm_data(host, username, password)

        except Exception as e:
            self.show_message_box(f"Error: {str(e)}")

    # Data fetch for cluster data (Registration in API)
    def fetch_cluster_data(self, host, username, password):
        host = self.cluster_ip
        username = self.username
        password = self.password

        if not host or not username or not password:
            self.show_message_box("Please enter your credentials in the Settings!")
            return
        
        try:
            url = f'https://{host}/rest/v1'
            credentials = 'Basic {0}'.format(str(base64.b64encode(bytes(f'{username}:{password}', 'utf-8')), 'utf-8'))
            rest_opts = {
                'Content-Type': 'application/json',
                'Authorization': credentials,
                'Connection': 'keep-alive'
            }
            
            context = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
            context.check_hostname = False
            context.verify_mode = ssl.CERT_NONE

            connection = http.HTTPSConnection(host, timeout=30, context=context)
            connection.request('GET', f'{url}/Registration', None, rest_opts)
            response = connection.getresponse()
            if response.status != 200:
                raise Exception(f"HTTP Error {response.status}")
            ClusterDataResult = json.loads(response.read().decode("utf-8"))
            connection.close()
            
            for xml in ClusterDataResult:
                xml_data = xml.get("clusterData")
                if xml_data:
                    # Parse the XML data (assuming it's a string)
                    root = ET.fromstring(xml_data)
                    
                    # Assuming that the XML structure contains tags and their corresponding values
                    cluster_data = {}
                    for elem in root.iter():
                        # Extract the tag name and the text (value) from each XML element
                        tag = elem.tag
                        value = elem.text.strip() if elem.text else "N/A"
                        cluster_data[tag] = value
                    
                    self.processed_cluster.update(cluster_data)

                    # Now, display this data in the treeview
                    for tag, value in cluster_data.items():
                        self.cluster_tree.insert("", ctk.END, values=(tag, value))

            self.alternate_row_colors()

        except Exception as e:
            self.show_message_box(f"Error fetching cluster data: {str(e)}")

    # Data fetch for VM data (VirDomain in API)
    def fetch_vm_data(self, host, username, password):
        host = self.cluster_ip
        username = self.username
        password = self.password
        
        try:
            url = f'https://{host}/rest/v1'
            credentials = 'Basic {0}'.format(str(base64.b64encode(bytes(f'{username}:{password}', 'utf-8')), 'utf-8'))
            rest_opts = {
                'Content-Type': 'application/json',
                'Authorization': credentials,
                'Connection': 'keep-alive'
            }
            
            context = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
            context.check_hostname = False
            context.verify_mode = ssl.CERT_NONE
            
            connection = http.HTTPSConnection(host, timeout=120, context=context)
            connection.request('GET', f'{url}/VirDomain', None, rest_opts)
            response = connection.getresponse()
            if response.status != 200:
                raise Exception(f"HTTP Error {response.status}")
            VirDomainResult = json.loads(response.read().decode("utf-8"))
            connection.close()
            
            for vm in VirDomainResult:
                memory = vm.get('mem')
                memory = round(memory / (1024 ** 3)) if isinstance(memory, (int, float)) else "N/A"
                vm_name = vm.get("name")
                if vm_name not in self.processed_vms:
                    self.processed_vms[vm_name] = {
                        "uuid": vm.get("uuid"),
                        "description": vm.get("description"),
                        "os": vm.get("operatingSystem", "N/A"),
                        "machineType": vm.get("machineType", "N/A"),
                        "state": vm.get("state", "N/A"),
                        "vcpus": vm.get("numVCPU", "N/A"),
                        "memory": memory,
                        "blocks": []
                    }
                
                # Nested data collection from blockDevs
                block_devices = vm.get("blockDevs", [])
                if not block_devices:
                    self.processed_vms[vm_name]["blocks"].append({
                        "name": "N/A",
                        "type": "N/A",
                        "capacity": "",
                        "allocation": "",
                        "mountPoints": "N/A"
                    })
                else:
                    for block in block_devices:
                        self.processed_vms[vm_name]["blocks"].append({
                            "name": f"{block.get('name', 'N/A')} ({block.get('uuid', 'N/A')})",
                            "type": block.get("type", "N/A"),
                            "capacity": round(float(block.get("capacity", 0)) / (1024 ** 3), 2) if block.get("capacity") else "",
                            "allocation": round(float(block.get("allocation", 0)) / (1024 ** 3), 2) if block.get("allocation") else "",
                            "mountPoints": block.get("mountPoints", "N/A")
                        })
            
            row_index = 0
            for vm_name, info in self.processed_vms.items():
                first_entry = True
                for block in info["blocks"]:
                    if block["type"] in ["NVRAM", "IDE_CDROM", "VTPM"]:
                        continue  
                    item_id = self.vm_tree.insert("", ctk.END, values=(
                        vm_name, 
                        info["uuid"] if first_entry else "",  
                        info["description"] if first_entry else "",
                        info["os"] if first_entry else "",
                        info["machineType"] if first_entry else "",
                        info["state"] if first_entry else "",
                        info["vcpus"] if first_entry else "",
                        info["memory"] if first_entry else "",
                        block["name"], 
                        block["type"], 
                        block["capacity"], 
                        block["allocation"],
                        block["mountPoints"]
                    ))
                    tag = 'evenrow' if row_index % 2 == 0 else 'oddrow'
                    self.vm_tree.item(item_id, tags=(tag,))
                    row_index += 1
                    first_entry = False

        except Exception as e:
            self.show_message_box(f"Error fetching VM data: {str(e)}")

        # Total calculations
        total_vcpus = 0
        total_memory = 0
        total_capacity = 0
        total_allocation = 0

        for vm_name, info in self.processed_vms.items():
            vcpus = str(info["vcpus"])
            total_vcpus += int(info["vcpus"]) if str(info["vcpus"]).isdigit() else 0
            total_memory += int(info["memory"]) if str(info["memory"]).isdigit() else 0
            
            for block in info["blocks"]:
                total_capacity += float(block["capacity"]) if block["capacity"] else 0
                total_allocation += float(block["allocation"]) if block["allocation"] else 0

        self.vm_tree.insert("", ctk.END, values=(
            "TOTAL", "", "", "", "", "", 
            total_vcpus, total_memory, 
            "", "", 
            round(total_capacity, 2), 
            round(total_allocation, 2),
            ""
        ), tags=('total'))

        self.vm_tree.tag_configure('total', background='#e3004b', foreground='white', font=("Martel Sans", 14))
        self.vm_tree.tag_configure('separator', background='#2e2e2e')

    # Export to Excel function
    def export(self):
        self.fetch_data(view_type="vm")
        self.fetch_data(view_type="cluster")

        if not self.processed_vms or not self.processed_cluster:
            self.show_message_box("Error: Insufficient data to export!")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All Files", "*.*")])
        if not file_path:
            return 

        # Prepare VM data
        vm_data = []
        total_vcpus = 0
        total_memory = 0
        total_capacity = 0
        total_allocation = 0

        for vm_name, info in self.processed_vms.items():
            for block in info["blocks"]:
                vm_data.append([
                    vm_name, 
                    info["uuid"], 
                    info["description"], 
                    info["os"], 
                    info["machineType"], 
                    info["state"], 
                    info["vcpus"], 
                    info["memory"], 
                    block["name"], 
                    block["type"], 
                    block["capacity"], 
                    block["allocation"],
                    block["mountPoints"]
                ])

            # Summing up values for the total line and float conversion
            total_vcpus += int(info["vcpus"]) if info["vcpus"] else 0
            total_memory += float(info["memory"]) if info["memory"] else 0
            total_capacity += float(block["capacity"]) if block["capacity"] else 0
            total_allocation += float(block["allocation"]) if block["allocation"] else 0

        # Append the total row
        vm_data.append([
            "TOTAL", "", "", "", "", "",
            total_vcpus, 
            total_memory, 
            "", "",  
            total_capacity, 
            total_allocation,
            ""
        ])

        vm_columns = ["Name", "UUID", "Description", "OS", "Machine Type", "State", 
                    "vCPUs", "Memory (GiB)", "Block Device", "Device Type", "Capacity (GiB)", 
                    "Allocation (GiB)", "Mount Points"]

        # Prepare Cluster Data
        cluster_data = []
        for tag, value in self.processed_cluster.items():
            cluster_data.append([tag, value])

        cluster_columns = ["Tag", "Value"]

        # Write to Excel
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            workbook = writer.book
        # Write Cluster sheet
            if cluster_data:
                df_cluster = pd.DataFrame(cluster_data, columns=cluster_columns)
                df_cluster.to_excel(writer, sheet_name="Cluster", index=False)
                worksheet_cluster = writer.sheets["Cluster"]

                # Auto-adjust column width for Cluster sheet
                for i, col in enumerate(cluster_columns):
                    max_length = max(df_cluster[col].astype(str).apply(len).max(), len(col)) + 2
                    worksheet_cluster.set_column(i, i, max_length)

            # Write Virtual Machines sheet
            if vm_data:
                df_vm = pd.DataFrame(vm_data, columns=vm_columns)
                df_vm.to_excel(writer, sheet_name="Virtual Machines", index=False)
                worksheet_vm = writer.sheets["Virtual Machines"]

                # Auto-adjust column width for Virtual Machines sheet
                for i, col in enumerate(vm_columns):
                    max_length = max(df_vm[col].astype(str).apply(len).max(), len(col)) + 2
                    worksheet_vm.set_column(i, i, max_length)

                # Apply formatting to the total row
                workbook = writer.book
                worksheet = writer.sheets["Virtual Machines"]
                bold_format = workbook.add_format({"bold": True})

                # Find the total row number
                total_row = len(df_vm) 

                # Apply bold formatting to the total row
                for col_num in range(len(vm_columns)):  
                    worksheet.write(total_row, col_num, df_vm.iloc[-1, col_num], bold_format)

        self.show_message_box("Successfully exported!")

if __name__ == "__main__":
    root = ctk.CTk()
    app = ClusterApp(root)
    root.mainloop()