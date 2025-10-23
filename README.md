# Pallet Packing Optimization Script

**Author:** Alexandre Cholat
**Date:** 2024
**Language:** Python

---

## Overview
This script optimizes the packing of items onto pallets for logistics and supply chain operations. It uses the **py3dbp** library to solve the 3D bin packing problem, ensuring efficient space utilization and weight distribution. Built for **industrial automation** and **smart logistics**, the tool integrates seamlessly with **cloud-based data pipelines** and **Excel/ERP systems** to fetch real-time order data. The tool is designed to:
- Automate pallet planning for orders.
- Optimize packing configurations to minimize wasted space.
- Reduce late deliveries by streamlining logistics.

The solution is **scalable**, **cloud-ready**, and can be deployed as part of a **warehouse management system (WMS)** or **supply chain execution (SCE) platform**, making it adaptable for both **on-premise** and **cloud-native** environments.


**Features:**
- **3D Visualization:** Generates a 3D model of packed pallets for visual validation.
- **Excel Integration:** Reads item data (dimensions, weight, quantity) from an Excel file.
- **Constraint Handling:** Accounts for load-bearing capacity, weight limits, and item dimensions.
- **Interactive GUI:** Displays packed pallets for user review before proceeding.

---

## How It Works
1. **Input:** Provide an order code (e.g., `1002128`).
2. **Processing:**
   - Fetches item details from an Excel file.
   - Calculates optimal packing using the **py3dbp** library.
   - Generates a 3D visualization of packed pallets.
3. **Output:** Returns the total number of pallets required for the order.

---

## Key Functions

| Function | Description |
|----------|-------------|
| `fetch_items_from_excel(command_code, excel_file)` | Reads item data from Excel and returns a list of `Item` objects. |
| `calculate_pallets(command_code)` | Computes the optimal number of pallets and visualizes the packing. |
| `visualize_pallets(packer)` | Renders a 3D plot of packed pallets using `matplotlib`. |
| `add_box(ax, item, color)` | Adds a 3D box (item) to the visualization. |

---

## Dependencies
Install the required libraries:
```bash
pip install py3dbp pandas matplotlib
