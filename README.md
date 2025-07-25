# 🗺️ Store Route Assignment Tool

A smart, GUI-based Python application to automatically assign uncovered stores to the nearest existing routes using geospatial distance and intelligent business rules.

---

## 📌 Features

- 🚀 **Automated Route Assignment**: Assigns unassigned stores to the nearest routes using geospatial matching (`BallTree` with Haversine distance).
- 🧠 **Rule Enforcement**: Supports prefix matching, branch matching, and route capacity limits (max 35 stores).
- 📊 **Excel Output**: Saves assigned results and summary report in a clean Excel workbook.
- 🖥️ **Modern GUI**: Built with `ttkbootstrap` for a dark-themed, responsive interface.
- 📍 **Custom Distance Radius**: Adjustable max distance threshold (in kilometers).
- 🧾 **Downloadable Templates**: Export sample Excel templates for covered and not-covered store data.

---

## 🛠️ Tech Stack

- `Python 3.x`
- `pandas`, `numpy`, `scikit-learn`
- `ttkbootstrap`, `tkinter`
- `openpyxl`

---

## 🚀 Getting Started

### 1. Clone this repository

```bash
git clone https://github.com/your-username/RouteAssignmentTool.git
cd RouteAssignmentTool
