import pandas as pd
import numpy as np
from sklearn.neighbors import BallTree
from collections import defaultdict
import ttkbootstrap as tb
from tkinter import filedialog, messagebox


def assign_routes(covered_df, not_covered_df, max_distance_km, enforce_limit=True, enforce_prefix=True, enforce_branch=True):
    covered_df = covered_df.copy()
    not_covered_df = not_covered_df.copy()

    covered_df['branch'] = covered_df['branch'].astype(str).str.lower()
    not_covered_df['branch'] = not_covered_df['branch'].astype(str).str.lower()

    covered_coords = np.radians(covered_df[['latitude', 'longitude']].values)
    not_covered_coords = np.radians(not_covered_df[['latitude', 'longitude']].values)

    tree = BallTree(covered_coords, metric='haversine')
    distances, indices = tree.query(not_covered_coords, k=10)
    distances_km = distances * 6371

    existing_counts = covered_df['route_code'].value_counts().to_dict()
    new_assignments = defaultdict(int)
    assignment_level_count = defaultdict(int)

    assigned_routes = []
    assigned_distances = []
    assigned_rank = []

    for i in range(len(not_covered_df)):
        assigned = False
        store_prefix = str(not_covered_df.iloc[i]['retailer_code'])[:5].lower()
        store_branch = not_covered_df.iloc[i]['branch']

        for j in range(5):
            idx = indices[i, j]
            route = covered_df.iloc[idx]['route_code']
            route_prefix = str(route)[:5].lower()
            route_branch = covered_df.iloc[idx]['branch']
            dist_km = distances_km[i, j]

            if dist_km > max_distance_km:
                continue
            if enforce_prefix and (store_prefix != route_prefix):
                continue
            if enforce_branch and (store_branch != route_branch):
                continue
            if enforce_limit:
                total = existing_counts.get(route, 0) + new_assignments[route]
                if total >= 33:
                    continue

            assigned_routes.append(route)
            assigned_distances.append(dist_km)
            new_assignments[route] += 1
            assignment_level_count[j + 1] += 1
            assigned_rank.append(j + 1)
            assigned = True
            break

        if not assigned:
            assigned_routes.append(None)
            assigned_distances.append(None)
            assigned_rank.append(None)

    not_covered_df['Assigned Route Code'] = assigned_routes
    not_covered_df['Distance_km'] = assigned_distances
    not_covered_df['Assignment Rank (1=nearest)'] = assigned_rank

    summary_data = {
        "Assignment Level": [f"{i} Nearest" for i in range(1, 6)],
        "Stores Assigned": [assignment_level_count.get(i, 0) for i in range(1, 6)],
    }
    unassigned = not_covered_df['Assigned Route Code'].isnull().sum()
    summary_df = pd.DataFrame(summary_data)
    summary_df = pd.concat([
        summary_df,
        pd.DataFrame([{"Assignment Level": "Unassigned", "Stores Assigned": unassigned}])
    ], ignore_index=True)

    return not_covered_df, summary_df


class RouteAssignmentApp:
    def __init__(self):
        self.root = tb.Window(themename="darkly")
        self.root.title("🗺️ Store Route Assignment Tool")
        self.root.geometry("800x520")

        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)

        frame = tb.Frame(self.root, padding=20)
        frame.grid(row=0, column=0, sticky="nsew")

        for i in range(9):
            frame.rowconfigure(i, weight=1)
        for j in range(3):
            frame.columnconfigure(j, weight=1)

        tb.Label(frame, text="📄 Covered Stores File:").grid(row=0, column=0, sticky='e', pady=6)
        self.covered_entry = tb.Entry(frame)
        self.covered_entry.grid(row=0, column=1, padx=8, sticky="ew")
        tb.Button(frame, text="Browse", command=self.browse_covered).grid(row=0, column=2, padx=8)

        tb.Label(frame, text="📄 Not Covered Stores File:").grid(row=1, column=0, sticky='e', pady=6)
        self.not_covered_entry = tb.Entry(frame)
        self.not_covered_entry.grid(row=1, column=1, padx=8, sticky="ew")
        tb.Button(frame, text="Browse", command=self.browse_not_covered).grid(row=1, column=2, padx=8)

        tb.Label(frame, text="💾 Output File Location:").grid(row=2, column=0, sticky='e', pady=6)
        self.output_entry = tb.Entry(frame)
        self.output_entry.grid(row=2, column=1, padx=8, sticky="ew")
        tb.Button(frame, text="Save As", command=self.save_output).grid(row=2, column=2, padx=8)

        tb.Label(frame, text="📏 Max Distance (km):").grid(row=3, column=0, sticky='e', pady=6)
        self.radius_entry = tb.Entry(frame, width=10)
        self.radius_entry.insert(0, "10")
        self.radius_entry.grid(row=3, column=1, sticky='w', padx=8)

        self.enforce_limit = tb.BooleanVar(value=True)
        tb.Checkbutton(frame, text="🔒 Enforce max 35 stores", variable=self.enforce_limit).grid(row=3, column=2, sticky='w')

        self.enforce_prefix = tb.BooleanVar(value=True)
        tb.Checkbutton(frame, text="🔤 Match prefix", variable=self.enforce_prefix).grid(row=4, column=2, sticky='w')

        self.enforce_branch = tb.BooleanVar(value=True)
        tb.Checkbutton(frame, text="🏢 Match branch", variable=self.enforce_branch).grid(row=5, column=2, sticky='w')

        tb.Button(frame, text="⬇️ Covered Template", command=self.download_covered_template).grid(row=4, column=0, pady=10)
        tb.Button(frame, text="⬇️ Not Covered Template", command=self.download_not_covered_template).grid(row=5, column=0)

        self.progress = tb.Progressbar(frame, length=400, mode="indeterminate")
        self.progress.grid(row=6, column=0, columnspan=3, pady=10)

        tb.Button(frame, text="🚀 Assign Routes", command=self.run_assignment).grid(row=7, column=0, columnspan=3, pady=10)
        tb.Label(frame, text="🔧 Developed by Abhijeet Kumar | Software Engineer", anchor="center", font=("Arial", 9), foreground="#aaa").grid(row=8, column=0, columnspan=3, pady=5)

    def browse_covered(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.covered_entry.delete(0, 'end')
            self.covered_entry.insert(0, path)

    def browse_not_covered(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.not_covered_entry.delete(0, 'end')
            self.not_covered_entry.insert(0, path)

    def save_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.output_entry.delete(0, 'end')
            self.output_entry.insert(0, path)

    def download_covered_template(self):
        df = pd.DataFrame(columns=['retailer_code', 'route_code', 'latitude', 'longitude', 'branch'])
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", title="Save Covered Template")
        if path:
            df.to_excel(path, index=False)
            messagebox.showinfo("Saved", "Covered Template saved successfully!")

    def download_not_covered_template(self):
        df = pd.DataFrame(columns=['retailer_code', 'latitude', 'longitude', 'branch'])
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", title="Save Not Covered Template")
        if path:
            df.to_excel(path, index=False)
            messagebox.showinfo("Saved", "Not Covered Template saved successfully!")

    def run_assignment(self):
        try:
            radius_km = float(self.radius_entry.get())
            enforce_limit = self.enforce_limit.get()
            enforce_prefix = self.enforce_prefix.get()
            enforce_branch = self.enforce_branch.get()

            covered_path = self.covered_entry.get()
            not_covered_path = self.not_covered_entry.get()
            output_path = self.output_entry.get()

            if not covered_path or not not_covered_path or not output_path:
                messagebox.showerror("Missing Input", "Please select all input files and output location.")
                return

            covered = pd.read_excel(covered_path)
            not_covered = pd.read_excel(not_covered_path)

            covered.columns = [col.lower() for col in covered.columns]
            not_covered.columns = [col.lower() for col in not_covered.columns]

            covered = covered.rename(columns={'retailercode': 'retailer_code'})
            not_covered = not_covered.rename(columns={'lattitude': 'latitude'})

            for df in [covered, not_covered]:
                df['latitude'] = pd.to_numeric(df['latitude'], errors='coerce')
                df['longitude'] = pd.to_numeric(df['longitude'], errors='coerce')
                df.dropna(subset=['latitude', 'longitude'], inplace=True)

            self.progress.start()

            assigned_df, summary_df = assign_routes(
                covered,
                not_covered,
                radius_km,
                enforce_limit=enforce_limit,
                enforce_prefix=enforce_prefix,
                enforce_branch=enforce_branch
            )

            with pd.ExcelWriter(output_path) as writer:
                assigned_df.to_excel(writer, sheet_name='Assigned Stores', index=False)
                summary_df.to_excel(writer, sheet_name='Assignment Summary', index=False)

            self.progress.stop()
            messagebox.showinfo("Success", "✅ Routes assigned and file saved successfully.")

        except Exception as e:
            self.progress.stop()
            messagebox.showerror("Error", f"Failed to assign routes:\n{e}")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = RouteAssignmentApp()
    app.run()
