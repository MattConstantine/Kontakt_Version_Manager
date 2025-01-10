''' 

Kontakt Version Manager

Written by Mr Thursday

Last modified 10/1/25

This script assumes that you keep all of your Kontakt exe files in the same folder named by version.
You could also include betas simply by adding a b at the end, or the word beta when storing

Select the mode 'load' to set the relevant exe and vst as the active version
Select the mode 'store' to save the currently active versions for future loading.
Select the mode 'read' to see what version is currently loaded.

Make sure that the correct version of Kontakt is selected.

'''

# ----------------------------------------------------------------------------------]
# Imports

from appdirs import user_config_dir
import tkinter as tk
from tkinter import ttk
import os
import shutil
from colorama import Fore, Style
import configparser
import webbrowser

# ----------------------------------------------------------------------------------]
# Options

APP_NAME             = 'Kontakt Version Manager'
APP_AUTHOR           = 'Its All Noise'     
CONFIG_SECTION       = "Settings"
VERSIONS_SECTION     = "Versions"

# ----------------------------------------------------------------------------------]
# Functions

def write_instructions(text_widget):
    instruction_text = '''Kontakt Version Manager
    
Set the Path to your Kontakt Library folder.  
This should contain all of your alternate Kontakt versions.

Kontakt Version (integer): Select the version of Kontakt you want to manage.
Kontakt Version (detail): Enter the specific version number you want to load or store.  
You can use any naming system.

include VST: Check this box if you want to manage the VST version of Kontakt.
include AAX: Check this box if you want to manage the AAX version of Kontakt.

Load: Set the selected version of Kontakt as the active version.

Store: Save the currently active version of Kontakt for future loading.

Read: Display the currently active version of Kontakt.

This program needs to be run with administrator privileges to access the Kontakt installation directories.
    '''
    config_file = get_config_path()
    final_text = f'{instruction_text}\n\nProgram Data Stored here:\n{config_file}'
    text_widget.insert(tk.END, final_text)

def get_config_path():
    """
    Returns the full path to a config.ini file in the user's OS-specific config directory.
    """
    config_dir = user_config_dir(appname=APP_NAME, appauthor=APP_AUTHOR)
    # Make sure the directory exists
    os.makedirs(config_dir, exist_ok=True)
    # Store settings in 'settings.ini' inside that directory
    return os.path.join(config_dir, "settings.ini")

def load_library_path():
    """
    Reads the 'LibraryPath' from the config file if it exists.
    """
    config_file = get_config_path()
    config = configparser.ConfigParser()

    if os.path.exists(config_file):
        config.read(config_file)
        if config.has_section(CONFIG_SECTION) and config.has_option(CONFIG_SECTION, "LibraryPath"):
            return config[CONFIG_SECTION]["LibraryPath"]
    return "Set the Library Path"

def save_library_path(path):
    """
    Saves 'LibraryPath' into the config file, creating the [Settings] section if necessary.
    """
    config_file = get_config_path()
    config = configparser.ConfigParser()

    # If a config file already exists, read it first to preserve other settings
    if os.path.exists(config_file):
        config.read(config_file)
    
    # Make sure section exists
    if not config.has_section(CONFIG_SECTION):
        config.add_section(CONFIG_SECTION)
    
    # Set the new path
    config[CONFIG_SECTION]["LibraryPath"] = path
    with open(config_file, "w", encoding="utf-8") as f:
        config.write(f)

def store_kontakt_version_in_config(kontakt_version, file_extension, version_string):
    """
    Stores the version for the specified kontakt_version and file_extension
    in the config file under a [Versions] section.
    """
    config_file = get_config_path()
    config = configparser.ConfigParser()
    
    if os.path.exists(config_file):
        config.read(config_file)
    
    if not config.has_section(VERSIONS_SECTION):
        config.add_section(VERSIONS_SECTION)
    
    # Use the kontakt_version and file_extension to create a key
    key = f"Kontakt {kontakt_version}{file_extension}"
    config[VERSIONS_SECTION][key] = f'Kontakt {version_string}{file_extension}'
    with open(config_file, "w", encoding="utf-8") as f:
        config.write(f)

def load_kontakt_version_from_config(kontakt_version, file_extension):
    """
    Returns the stored version string for the given kontakt_version and file_extension,
    """
    config_file = get_config_path()
    config = configparser.ConfigParser()

    if os.path.exists(config_file):
        config.read(config_file)
        if config.has_section(VERSIONS_SECTION):
            key = f"Kontakt {kontakt_version}{file_extension}"
            if key in config[VERSIONS_SECTION]:
                return config[VERSIONS_SECTION][key]
    nil_message = f"You need to Load or Store a Kontakt Version for Kontakt {kontakt_version}{file_extension} before reading is possible"
    return nil_message

def check_version_match(new_version,kontakt_version,text_widget):
    # Extract the first part of the version number before any space or dot
    first_part = new_version.split()[0].split('.')[0]
    
    try:
        # Convert the extracted part to an integer
        version_number = int(first_part)
    except ValueError:
        print(f"{Fore.RED}Error: Invalid version format.{Style.RESET_ALL}")
        text_widget.insert(tk.END,f"\nError: Invalid version format.")
        match = False
        return match
    
    # Compare the extracted version number with kontakt_version
    if version_number != kontakt_version:
        print(f"{Fore.RED}Kontakt Version mismatch{Style.RESET_ALL}")
        text_widget.insert(tk.END,"\nKontakt Version mismatch")
        match = False
        return match
    else:
        match = True
        return match

def copy_kontakt(source, destination, file_extension,new_version,kontakt_version,text_widget):
    """
    Loads Selected version as current working version of Kontakt.
    """
    if not os.path.isdir(source):
        print(f"{Fore.RED}Directory does not exist or cannot be accessed : \n{source}{Style.RESET_ALL}")
        text_widget.insert(tk.END,f"\nDirectory does not exist or cannot be accessed : \n{source}")
        return

    file_to_copy = f'Kontakt {new_version}{file_extension}'
    source_file_path = os.path.join(source, file_to_copy)

    if not os.path.exists(source_file_path):
        print(f"{Fore.RED}{file_to_copy} is not available.{Fore.YELLOW} \nAvailable versions include : {Style.RESET_ALL}")
        text_widget.insert(tk.END,f"\n{file_to_copy} is not available.\nAvailable versions include :")
        # List all files that start with "Kontakt" and end with the correct file extension
        for file in os.listdir(source):
            if file.startswith(f"Kontakt {kontakt_version}") and file.endswith(f"{file_extension}"):
                print(file)
                text_widget.insert(tk.END,f"\n{file}")
        print("\n")
        text_widget.insert(tk.END,"\n")
        return

    # Copy it directly to the destination, overwriting the file already there
    try:
        shutil.copy2(source_file_path, destination)
        print(f"{Fore.GREEN}{file_to_copy} copied to :\n{Fore.CYAN}{destination}{Style.RESET_ALL}\n")
        text_widget.insert(tk.END,f"\n{file_to_copy} loaded")
        store_kontakt_version_in_config(kontakt_version, file_extension, new_version)
    except Exception as e:
        print(f"{Fore.RED}Failed to copy {file_to_copy} to {destination}. Error: {e}{Style.RESET_ALL}\n")
        text_widget.insert(tk.END,f"\nFailed to load {file_to_copy}. Error: {e}\n")

def store_kontakt(source, destination,new_version,kontakt_version,text_widget):
    """
    Stores current working version of Kontakt into the version library.
    """
    filename,file_extension  = os.path.splitext(source)
    file_to_store  = f'Kontakt {new_version}{file_extension}'
    dest_file_path = os.path.join(destination, file_to_store)

    # Check if the specific version file exists in the source directory
    if not os.path.exists(source):
        print(f"{Fore.RED}{filename}{file_extension} not found in the source directory: \n{source}{Style.RESET_ALL}\n")
        text_widget.insert(tk.END,f"\n{filename}{file_extension} not found")
        return

    # Check if the destination file already exists to avoid overwriting
    if os.path.exists(dest_file_path):
        print(f"{Fore.YELLOW}{file_to_store} already exists and will not be overwritten: \n{Fore.CYAN}{dest_file_path}{Style.RESET_ALL}\n")
        text_widget.insert(tk.END,f"\n{file_to_store} already exists and will not be overwritten")
        return
    
    # Copy it directly to the destination, but only if it doesn't already exist
    try:
        shutil.copy2(source, dest_file_path)
        print(f"{Fore.GREEN}{file_to_store} copied to :\n{Fore.CYAN}{dest_file_path}{Style.RESET_ALL}\n")
        text_widget.insert(tk.END,f"\n{file_to_store} Stored")
        store_kontakt_version_in_config(kontakt_version, file_extension, new_version)
    except Exception as e:
        print(f"{Fore.RED}Failed to copy {file_to_store} file to \n{destination}. \nError: {e}{Style.RESET_ALL}\n")
        text_widget.insert(tk.END,f"\nFailed to store {file_to_store}. \nError: {e}\n")

def read_kontakt_version(source,kontakt_version,text_widget):
    """
    Checks the config file to see the current loaded version and print it.
    """
    if not os.path.exists(source):
        print(f"{Fore.RED}File not found : {source}{Style.RESET_ALL}\n")
        return

    _, file_extension = os.path.splitext(source)
    stored_version = load_kontakt_version_from_config(kontakt_version, file_extension)
    print(f"{Fore.GREEN}{stored_version}{Style.RESET_ALL}\n")
    text_widget.insert(tk.END, f"\n{stored_version}")

def run_kontakt_operation(mode, kontakt_version, new_version, include_vst, include_aax,text_widget,library_path):
    """
    Runs the essential logic for the selected operation mode.
    """
    kontakt_exe_path,kontakt_vst_path,kontakt_aax_path = set_kontakt_version(kontakt_version)

    if mode == 'read':
        print("Currently Loaded Versions:")
        text_widget.insert(tk.END,"\nCurrently Loaded Versions:")
        read_kontakt_version(kontakt_exe_path,kontakt_version,text_widget)
        if include_vst:
            read_kontakt_version(kontakt_vst_path,kontakt_version,text_widget)
        if include_aax:
            read_kontakt_version(kontakt_aax_path,kontakt_version,text_widget)
        text_widget.insert(tk.END,f"\n\nThis display shows the version last loaded or stored.  This could differ from the actual version present if the version has been changed by other methods")
    else:
        match = check_version_match(new_version,kontakt_version,text_widget)
        if match:
            if mode == 'load':
                copy_kontakt(library_path, kontakt_exe_path, '.exe',new_version,kontakt_version,text_widget)
                if include_vst:
                    copy_kontakt(library_path, kontakt_vst_path, '.vst3',new_version,kontakt_version,text_widget)
                if include_aax:
                    copy_kontakt(library_path, kontakt_aax_path, '.aaxplugin',new_version,kontakt_version,text_widget)
                text_widget.see(tk.END)
            elif mode == 'store':
                store_kontakt(kontakt_exe_path, library_path,new_version,kontakt_version,text_widget)
                if include_vst:
                    store_kontakt(kontakt_vst_path, library_path,new_version,kontakt_version,text_widget)
                if include_aax:
                    store_kontakt(kontakt_aax_path, library_path,new_version,kontakt_version,text_widget)

def set_kontakt_version(kontakt_version):
    """
    Sets the file paths - this is windows specific.
    """
    if kontakt_version   == 8: 
        kontakt_exe_path = 'C:\\Program Files\\Native Instruments\\Kontakt 8\\Kontakt 8.exe'
        kontakt_vst_path = 'C:\\Program Files\\Common Files\\VST3\\Kontakt 8.vst3'
        kontakt_aax_path = 'C:\\Program Files\\Common Files\\Avid\\Audio\\Plug-Ins\\Kontakt 8.aaxplugin\\Contents\\x64\\Kontakt 8.aaxplugin'
    elif kontakt_version == 7: 
        kontakt_exe_path = 'C:\\Program Files\\Native Instruments\\Kontakt 7\\Kontakt 7.exe'
        kontakt_vst_path = 'C:\\Program Files\\Common Files\\VST3\\Kontakt 7.vst3'
        kontakt_aax_path = 'C:\\Program Files\\Common Files\\Avid\\Audio\\Plug-Ins\\Kontakt 7.aaxplugin\\Contents\\x64\\Kontakt 7.aaxplugin'
    elif kontakt_version == 6: 
        kontakt_exe_path = 'C:\\Program Files\\Native Instruments\\Kontakt\\Kontakt.exe'
        kontakt_vst_path = 'C:\\Program Files\\Common Files\\VST3\\Kontakt.vst3'
        kontakt_aax_path = 'C:\\Program Files\\Common Files\\Avid\\Audio\\Plug-Ins\\Kontakt.aaxplugin\\Contents\\x64\\Kontakt.aaxplugin'
    elif kontakt_version == 5: 
        kontakt_exe_path = 'C:\\Program Files\\Native Instruments\\Kontakt 5\\Kontakt 5.exe'
        kontakt_vst_path = 'C:\\Program Files\\Steinberg\\VSTPlugins\\Native Instruments64\\Kontakt 5.dll'
        kontakt_aax_path = 'C:\\Program Files\\Common Files\\Avid\\Audio\\Plug-Ins\\Kontakt 5.aaxplugin\\Contents\\x64\\Kontakt 5.aaxplugin'
        
    return kontakt_exe_path,kontakt_vst_path,kontakt_aax_path

def main():
    # Create the main window
    root = tk.Tk()
    root.title("Kontakt Version Manager")
    root.geometry("700x700")
    
    # the library path
    tk.Label(root, text="Path to version Library (storing alternate versions)").pack(anchor="w", padx=30,pady=5)
    # Frame to hold the path entry and browse button
    path_frame = ttk.Frame(root)
    path_frame.pack(anchor="w", padx=30)
    library_path_var = tk.StringVar()

    # Load last-used path from config
    default_path = load_library_path()
    library_path_var.set(default_path)
    
    # Browse button
    def on_browse():
        from tkinter import filedialog
        initial_dir = library_path_var.get() or ""
        selected = filedialog.askdirectory(initialdir=initial_dir, title="Select Library Folder")
        if selected:
            library_path_var.set(selected)
            # Save to config each time user picks a new folder
            save_library_path(selected)
            
    browse_button = ttk.Button(path_frame, text="Browse", command=on_browse)
    browse_button.pack(side="left")
    
    # Display it in an Entry
    path_entry = ttk.Entry(path_frame, textvariable=library_path_var, width=80)
    path_entry.pack(side="left", padx=2)

    # Kontakt Version
    main_version_frame = ttk.Frame(root)
    main_version_frame.pack(anchor="w", padx=30, pady=10)    
    tk.Label(main_version_frame, text="Kontakt Version (integer):").pack(side="left")
    version_var = tk.StringVar(value="8")
    version_entry = ttk.Combobox(
        main_version_frame,
        textvariable=version_var,
        values=["5", "6", "7","8"],
        state="readonly",
        width=10
    )
    version_entry.pack(side="left", padx=5)

    # New Version
    detail_version_frame = ttk.Frame(root)
    detail_version_frame.pack(anchor="w", padx=30, pady=10)
    tk.Label(detail_version_frame, text="Kontakt Version (detail) :").pack(side="left")
    new_version_var = tk.StringVar(value="8.0.0")
    new_version_entry = ttk.Entry(detail_version_frame, textvariable=new_version_var, width=20)
    new_version_entry.pack(side="left")

    # Include VST
    include_vst_var = tk.BooleanVar(value=True)
    vst_check = ttk.Checkbutton(root, text="Include VST", variable=include_vst_var)
    vst_check.pack(anchor="w", padx=30)

    # Include AAX
    include_aax_var = tk.BooleanVar(value=True)
    aax_check = ttk.Checkbutton(root, text="Include AAX", variable=include_aax_var)
    aax_check.pack(anchor="w", padx=30)
    
    # Status label for a quick summary of the operation
    status_var = tk.StringVar(value="")
    status_label = ttk.Label(root, textvariable=status_var, wraplength=300)
    status_label.pack(pady=10)
    
    # ------------------
    # Define Callback Functions for each button: Load, Store, Read
    def on_load():
        text_output.delete("1.0", tk.END)  # clear previous output

        try:
            kontakt_version_int = int(version_var.get())
        except ValueError:
            status_var.set("Error: Kontakt version must be an integer.")
            return

        new_version = new_version_var.get()
        include_vst = include_vst_var.get()
        include_aax = include_aax_var.get()

        # Call the "load" operation
        run_kontakt_operation(
            mode="load",
            kontakt_version=kontakt_version_int,
            new_version=new_version,
            include_vst=include_vst,
            include_aax=include_aax,
            text_widget=text_output,
            library_path=library_path_var.get()
        )
        status_var.set("Completed load operation!")

    def on_store():
        text_output.delete("1.0", tk.END)

        try:
            kontakt_version_int = int(version_var.get())
        except ValueError:
            status_var.set("Error: Kontakt version must be an integer.")
            return

        new_version = new_version_var.get()
        include_vst = include_vst_var.get()
        include_aax = include_aax_var.get()

        # Call the "store" operation
        run_kontakt_operation(
            mode="store",
            kontakt_version=kontakt_version_int,
            new_version=new_version,
            include_vst=include_vst,
            include_aax=include_aax,
            text_widget=text_output,
            library_path=library_path_var.get()
        )
        status_var.set("Completed store operation!")

    def on_read():
        text_output.delete("1.0", tk.END)

        try:
            kontakt_version_int = int(version_var.get())
        except ValueError:
            status_var.set("Error: Kontakt version must be an integer.")
            return

        new_version = new_version_var.get()
        include_vst = include_vst_var.get()
        include_aax = include_aax_var.get()

        # Call the "read" operation
        run_kontakt_operation(
            mode="read",
            kontakt_version=kontakt_version_int,
            new_version=new_version,
            include_vst=include_vst,
            include_aax=include_aax,
            text_widget=text_output,
            library_path=library_path_var.get()
        )

        status_var.set(f"Completed read operation!")

    # ------------------
    
    # Frame to hold the three buttons horizontally
    btn_frame = ttk.Frame(root)
    btn_frame.pack(pady=10)
    btn_load = ttk.Button(btn_frame, text="Load", command=on_load)
    btn_load.pack(side="left", padx=5)
    btn_store = ttk.Button(btn_frame, text="Store", command=on_store)
    btn_store.pack(side="left", padx=5)
    btn_read = ttk.Button(btn_frame, text="Read", command=on_read)
    btn_read.pack(side="left", padx=5)

    #credit link
    link_label = tk.Label(
        root, 
        text="Made by Matt Constantine : https://mattconstantine.co.uk",
        fg="blue",            # text color
        cursor="hand2"        # 'hand2' is the typical link cursor in Tkinter
    )
    link_label.pack(pady=5)
    def open_mattconstantine_website(event):
        webbrowser.open("https://mattconstantine.co.uk")
    link_label.bind("<Button-1>", open_mattconstantine_website)
    
    # Main Feedback Window
    text_output = tk.Text(root, wrap='word', height=40, width=100)
    text_output.pack(pady=10)

    #display instructions at startup
    write_instructions(text_output)
    # Start the event loop
    root.mainloop()

# ----------------------------------------------------------------------------------]
# Execute script

if __name__ == "__main__":
    main()
