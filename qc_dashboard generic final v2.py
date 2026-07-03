print("PYTHON IS AWAKE! Testing imports one by one...")
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import numpy as np
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

print("ALL IMPORTS SUCCESSFUL! Launching app...\n")

# ==========================================
# 1. DYNAMIC FILE SELECTION 
# ==========================================

def select_files():
    print("[DEBUG] Step 1: Starting file selection process...")
    try:
        temp_root = tk.Tk()
        temp_root.title("Loading...")
        temp_root.geometry("250x100") 
        temp_root.attributes('-topmost', True)
        temp_root.lift()
        temp_root.focus_force()
        
        messagebox.showinfo("Select Data File", "Please select your QC Data Excel file.")
        
        data_file = filedialog.askopenfilename(
            title="Select QC Data File", 
            filetypes=[("Excel files", "*.xlsx *.xlsm")]
        )
        print(f"[DEBUG] Step 2: Data file selected: {data_file}")
        
        if data_file: 
            messagebox.showinfo("Select Limits File", "Please select your Limits/Targets Excel file.")
            targets_file = filedialog.askopenfilename(
                title="Select Limits/Targets File", 
                filetypes=[("Excel files", "*.xlsx *.xlsm")]
            )
            print(f"[DEBUG] Step 3: Targets file selected: {targets_file}")
        else:
            targets_file = ""
            
        temp_root.destroy() 
        
        if not data_file or not targets_file:
            return None, None
            
        return data_file, targets_file
        
    except Exception as e:
        print(f"[CRITICAL ERROR] Failed during file selection: {e}")
        return None, None

# ==========================================
# 2. DATA LOADING & CLEANING (ULTRA-FAST)
# ==========================================
def load_and_clean_data(file_path):
    """Loads and cleans data dynamically."""
    print("[DEBUG] Step 4: Loading Data File into Pandas...")
    try:
        df = pd.read_excel(file_path, sheet_name="QC_Data", engine="calamine")
    except ValueError:
        df = pd.read_excel(file_path, engine="calamine")
    except Exception as e:
        df = pd.read_excel(file_path)

    df.columns.values[0] = "Batch ID"
    df.columns.values[1] = "PT Round"
    
    df = df.dropna(subset=["PT Round"])
    df["PT Round"] = df["PT Round"].astype(str).str.strip()
    
    df = df[~df["PT Round"].isin(["nan", "NaN", "0", "0.0", ""])]
    
    cols_to_convert = df.columns[2:]
    df[cols_to_convert] = df[cols_to_convert].apply(pd.to_numeric, errors='coerce')

    return df

def load_targets(file_path):
    """Loads targets using Calamine for extreme speed."""
    print("[DEBUG] Step 5: Loading Limits File into Pandas...")
    try:
        df_targets = pd.read_excel(file_path, sheet_name="Sheet1", engine="calamine")
    except ValueError:
        df_targets = pd.read_excel(file_path, engine="calamine")
    except Exception:
        df_targets = pd.read_excel(file_path)

    df_targets.columns = df_targets.columns.astype(str).str.strip() 
    df_targets["PT Round"] = df_targets["PT Round"].astype(str).str.strip()
    df_targets["Analyte"] = df_targets["Analyte"].astype(str).str.strip()
    
    # Setting the index strips the first two columns out of the normal 0-index positional counting
    df_targets.set_index(["PT Round", "Analyte"], inplace=True)
    
    return df_targets

def get_limits(targets_df, pt_round, analyte):
    """Extracts analytical targets, bounds, Matrix Type, and Columns H & I logic."""
    try:
        target_data = targets_df.loc[(pt_round, analyte)]
        
        if isinstance(target_data, pd.DataFrame):
            target_data = target_data.iloc[0]
            
        # Pos 0 = Col C (Assigned Value)
        mean_val = pd.to_numeric(target_data.iloc[0], errors='coerce')
        # Pos 1 = Col D (Upper Limit / +2SD)
        warn_upper = pd.to_numeric(target_data.iloc[1], errors='coerce')
        # Pos 2 = Col E (Lower Limit / -2SD)
        warn_lower = pd.to_numeric(target_data.iloc[2], errors='coerce')
        
        matrix = str(target_data.iloc[4]).strip() if len(target_data) > 4 else "Unknown"

        # Pos 5 = Col H (e.g., 1.2 or +2SD), Pos 6 = Col I (e.g., 0.8 or -2SD)
        col_H = str(target_data.iloc[5]).strip() if len(target_data) > 5 else "+2SD"
        col_I = str(target_data.iloc[6]).strip() if len(target_data) > 6 else "-2SD"
            
        return mean_val, warn_upper, warn_lower, matrix, col_H, col_I
    except Exception as e:
        print(f"[DEBUG] Target Error: {e}")
        return np.nan, np.nan, np.nan, "Unknown", "+2SD", "-2SD"

# ==========================================
# 3. PROFESSIONAL NATIVE DESKTOP GUI
# ==========================================
def launch_desktop_app():
    data_file, targets_file = select_files()
    
    if not data_file or not targets_file:
        print("[DEBUG] File selection cancelled. Exiting.")
        return
        
    df = load_and_clean_data(data_file)
    targets_df = load_targets(targets_file)
    print("[DEBUG] Step 6: Data loaded successfully. Drawing main UI...")
    
    # --- UI SYSTEM CONSTANTS ---
    BG_SIDEBAR = "#F3F6F9"
    BG_MAIN = "#FFFFFF"
    TEXT_MAIN = "#2C3E50"
    TEXT_SUB = "#7F8C8D"
    ACCENT = "#0078D7"
    FONT_FAMILY = "Segoe UI"
    
    mpl.rcParams['font.family'] = FONT_FAMILY
    mpl.rcParams['text.color'] = TEXT_MAIN
    mpl.rcParams['axes.labelcolor'] = TEXT_SUB
    mpl.rcParams['xtick.color'] = TEXT_SUB
    mpl.rcParams['ytick.color'] = TEXT_SUB

    root = tk.Tk()
    root.title("OEC Laboratory | QC Dashboard")
    root.geometry("1200x750") 
    root.configure(bg=BG_MAIN)
    
    style = ttk.Style()
    if 'clam' in style.theme_names():
        style.theme_use('clam')
        
    style.configure("Sidebar.TFrame", background=BG_SIDEBAR)
    style.configure("Header.TLabel", background=BG_SIDEBAR, foreground=TEXT_MAIN, font=(FONT_FAMILY, 16, "bold"))
    style.configure("Sub.TLabel", background=BG_SIDEBAR, foreground=TEXT_SUB, font=(FONT_FAMILY, 10))
    style.configure("Title.TLabel", background=BG_SIDEBAR, foreground=TEXT_MAIN, font=(FONT_FAMILY, 11, "bold"))

    # --- LAYOUT SEPARATORS ---
    left_frame = ttk.Frame(root, width=320, style="Sidebar.TFrame")
    left_frame.pack(side="left", fill="y")
    left_frame.pack_propagate(False) 
    
    sidebar_pad = tk.Frame(left_frame, bg=BG_SIDEBAR, padx=25, pady=25)
    sidebar_pad.pack(fill="both", expand=True)
    
    right_frame = tk.Frame(root, bg=BG_MAIN)
    right_frame.pack(side="right", fill="both", expand=True)

    # --- CONTROL ELEMENTS ---
    ttk.Label(sidebar_pad, text="Laboratory QC", style="Header.TLabel").pack(anchor="w", pady=(0, 2))
    
    filename = os.path.basename(data_file)
    try:
        date_part = filename.split('_')[-1].split('.')[0]
        if len(date_part) == 8 and date_part.isdigit():
            date_text = f"Data Updated: {date_part[:4]}/{date_part[4:6]}/{date_part[6:]}"
        else:
            date_text = f"Data Updated: {date_part}"
    except:
        date_text = "Latest Data Loaded"
        
    ttk.Label(sidebar_pad, text=date_text, style="Sub.TLabel").pack(anchor="w", pady=(0, 30))
    
    ttk.Label(sidebar_pad, text="1. Select PT Round", style="Title.TLabel").pack(anchor="w")
    available_pt_rounds = list(df["PT Round"].unique())
    if "T06133" in available_pt_rounds:
        available_pt_rounds.remove("T06133")
        available_pt_rounds.insert(0, "T06133")
        
    pt_var = tk.StringVar()
    pt_dropdown = ttk.Combobox(sidebar_pad, textvariable=pt_var, values=available_pt_rounds, state="readonly", font=(FONT_FAMILY, 10))
    pt_dropdown.pack(fill="x", pady=(5, 25))
    if available_pt_rounds:
        pt_dropdown.current(0)

    ttk.Label(sidebar_pad, text="2. Select Analyte", style="Title.TLabel").pack(anchor="w")
    list_frame = tk.Frame(sidebar_pad, bg=BG_SIDEBAR)
    list_frame.pack(fill="x", pady=(5, 25))
    
    scrollbar = tk.Scrollbar(list_frame, orient="vertical")
    scrollbar.pack(side="right", fill="y")
    
    analyte_listbox = tk.Listbox(list_frame, selectmode="single", yscrollcommand=scrollbar.set, 
                                 height=14, font=(FONT_FAMILY, 10), 
                                 bg="#FFFFFF", fg=TEXT_MAIN, 
                                 selectbackground=ACCENT, selectforeground="white",
                                 relief="flat", highlightthickness=1, highlightcolor="#CBD5E1", highlightbackground="#CBD5E1")
    analyte_listbox.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=analyte_listbox.yview)

    ttk.Label(sidebar_pad, text="3. Chart View Range", style="Title.TLabel").pack(anchor="w")
    view_var = tk.StringVar()
    view_dropdown = ttk.Combobox(sidebar_pad, textvariable=view_var, values=["Last 4 points", "Last 25 points", "Last 50 points", "Last 100 points", "All Data"], state="readonly", font=(FONT_FAMILY, 10))
    view_dropdown.pack(fill="x", pady=(5, 20))
    view_dropdown.set("Last 50 points")

    # --- CHART SYSTEM PLOT ---
    fig, ax = plt.subplots(figsize=(9, 6), dpi=100)
    fig.patch.set_facecolor(BG_MAIN)
    ax.set_facecolor(BG_MAIN)
    
    canvas = FigureCanvasTkAgg(fig, master=right_frame)
    canvas.get_tk_widget().pack(side="top", fill="both", expand=True, padx=20, pady=(20, 0))
    
    toolbar = NavigationToolbar2Tk(canvas, right_frame)
    toolbar.config(background=BG_MAIN)
    toolbar._message_label.config(background=BG_MAIN, font=(FONT_FAMILY, 9))
    toolbar.update()
    canvas.get_tk_widget().pack(side="top", fill="both", expand=True)

    def draw_chart(event=None):
        selected_pt = pt_var.get()
        selected_indices = analyte_listbox.curselection()
        if not selected_indices:
            return
            
        selected_analyte = analyte_listbox.get(selected_indices[0])
            
        filtered_df = df[df["PT Round"] == selected_pt]
        clean_df = filtered_df.dropna(subset=[selected_analyte])
        clean_df = clean_df[clean_df[selected_analyte] > 0]
        
        batch_ids = clean_df["Batch ID"].astype(str).tolist()
        y_values = clean_df[selected_analyte].values
        total_points = len(batch_ids)
        
        mean_val, warn_upper, warn_lower, matrix_type, col_H, col_I = get_limits(targets_df, selected_pt, selected_analyte)
        has_limits = pd.notna(mean_val)
        
        # ---------------------------------------------------------
        # LIMITS LOGIC: Switch between +/- 2SD and Percentage Mode
        # ---------------------------------------------------------
        if has_limits:
            col_H_str = str(col_H).strip().upper()
            
            # MODE A: Standard SD
            if "SD" in col_H_str or col_H_str == "NAN":
                sd_val = abs(warn_upper - mean_val) / 2
                action_upper = mean_val + (3 * sd_val)
                action_lower = mean_val - (3 * sd_val)
                
                lbl_warn_u = f"+2 SD (Warning): {warn_upper:.2f}"
                lbl_warn_d = f"-2 SD (Warning): {warn_lower:.2f}"
                lbl_act_u = f"+3 SD (Action): {action_upper:.2f}"
                lbl_act_d = f"-3 SD (Action): {action_lower:.2f}"
                
            # MODE B: Percentage Bounds (e.g. 1.2 and 0.8)
            else:
                try:
                    u_mult = float(col_H)  # e.g., 1.2 (120%)
                    l_mult = float(col_I)  # e.g., 0.8 (80%)
                    
                    # Estimate action limits proportionally (e.g. 130% / 70%)
                    dist_u = u_mult - 1.0
                    dist_l = 1.0 - l_mult
                    action_upper = mean_val * (1.0 + dist_u * 1.5)
                    action_lower = mean_val * (1.0 - dist_l * 1.5)
                    
                    lbl_warn_u = f"{int(u_mult*100)}% Limit (Warning): {warn_upper:.2f}"
                    lbl_warn_d = f"{int(l_mult*100)}% Limit (Warning): {warn_lower:.2f}"
                    lbl_act_u = f"{int((1.0 + dist_u*1.5)*100)}% Limit (Action): {action_upper:.2f}"
                    lbl_act_d = f"{int((1.0 - dist_l*1.5)*100)}% Limit (Action): {action_lower:.2f}"
                except ValueError:
                    # Failsafe fallback if conversion fails
                    sd_val = abs(warn_upper - mean_val) / 2
                    action_upper = mean_val + (3 * sd_val)
                    action_lower = mean_val - (3 * sd_val)
                    lbl_warn_u = f"Upper Limit: {warn_upper:.2f}"
                    lbl_warn_d = f"Lower Limit: {warn_lower:.2f}"
                    lbl_act_u = f"Action Upper: {action_upper:.2f}"
                    lbl_act_d = f"Action Lower: {action_lower:.2f}"

            colors = []
            for y in y_values:
                if y > action_upper or y < action_lower:
                    colors.append("#D32F2F")  # Red for Action
                elif y > warn_upper or y < warn_lower:
                    colors.append("#F57C00")  # Orange for Warning
                else:
                    colors.append("#1E88E5")  # Blue for Pass
        else:
            colors = ['#34495E'] * total_points

        ax.clear()

        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_color('#CBD5E1')
        ax.spines['bottom'].set_color('#CBD5E1')

        x_indices = np.arange(total_points)
        ax.plot(x_indices, y_values, linestyle='-', color='#B0BEC5', linewidth=1.5, zorder=1) 
        if len(x_indices) > 0: 
            ax.scatter(x_indices, y_values, c=colors, s=55, zorder=3, edgecolors='white', linewidth=0.5)

        ax.set_xticks(x_indices)
        ax.set_xticklabels(batch_ids, rotation=45, ha='right', fontsize=9) 

        if has_limits:
            ax.axhspan(warn_lower, warn_upper, color='#E8F5E9', alpha=0.5, zorder=0)  
            ax.axhspan(warn_upper, action_upper, color='#FFF3E0', alpha=0.5, zorder=0)    
            ax.axhspan(action_lower, warn_lower, color='#FFF3E0', alpha=0.5, zorder=0)  
            
            ax.axhline(y=mean_val, color='#2E7D32', linestyle='-', linewidth=2, label=f"Target (Mean): {mean_val:.2f}")
            ax.axhline(y=warn_upper, color='#F57C00', linestyle='--', linewidth=1.5, label=lbl_warn_u)
            ax.axhline(y=warn_lower, color='#F57C00', linestyle='--', linewidth=1.5, label=lbl_warn_d)
            ax.axhline(y=action_upper, color='#D32F2F', linestyle=':', linewidth=1.5, label=lbl_act_u)
            ax.axhline(y=action_lower, color='#D32F2F', linestyle=':', linewidth=1.5, label=lbl_act_d)
            
            ax.legend(loc='upper left', bbox_to_anchor=(1.02, 1), frameon=False, fontsize=9)

        ax.set_title(f"Levey-Jennings QC Chart: {selected_analyte}\nPT Round: {selected_pt} | Matrix: {matrix_type}", 
                     fontsize=14, fontweight='bold', color=TEXT_MAIN, pad=15, loc='left')
        
        ax.set_ylabel("Concentration", fontsize=10, labelpad=10)
        ax.grid(axis='y', linestyle='--', alpha=0.4, color='#CBD5E1')
        ax.grid(axis='x', visible=False)

        view_selection = view_var.get()
        if view_selection == "Last 4 points" and total_points > 4:
            ax.set_xlim(total_points - 4.5, total_points - 0.5)
        elif view_selection == "Last 25 points" and total_points > 25:
            ax.set_xlim(total_points - 25.5, total_points - 0.5)
        elif view_selection == "Last 50 points" and total_points > 50:
            ax.set_xlim(total_points - 50.5, total_points - 0.5)
        elif view_selection == "Last 100 points" and total_points > 100:
            ax.set_xlim(total_points - 100.5, total_points - 0.5)
        else:
            ax.set_xlim(-0.5, total_points - 0.5)

        fig.tight_layout()
        canvas.draw()

    def update_analytes(*args):
        selected_pt = pt_var.get()
        targets_flat = targets_df.reset_index()
        valid_analytes = list(targets_flat[targets_flat["PT Round"] == selected_pt]["Analyte"].unique())
        
        data_columns = df.columns[2:]
        final_list = [a for a in data_columns if a in valid_analytes]
        
        analyte_listbox.delete(0, tk.END)
        for item in final_list:
            analyte_listbox.insert(tk.END, item)
            
        if final_list:
            analyte_listbox.selection_set(0)
            draw_chart() 

    pt_var.trace_add("write", update_analytes)
    analyte_listbox.bind('<<ListboxSelect>>', draw_chart)
    view_var.trace_add("write", lambda *args: draw_chart()) 
    
    update_analytes() 
    
    def on_closing():
        root.quit()     
        root.destroy()  
        
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    print("[DEBUG] Step 7: Starting Main Loop! (UI should be visible now)")
    root.mainloop()

if __name__ == "__main__":
    launch_desktop_app()
