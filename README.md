# SAP‑Connected Excel Execution Engine
*A governed Excel‑based execution system for synchronising SAP Business One operational data.*

---

## Overview

This repository contains an Excel‑based execution system that connects directly to **SAP Business One’s SQL Server** and refreshes multiple operational views in a **single governed run**.

It was originally designed to support **inventory, purchasing, and logistics teams** who relied on Excel as their primary operational interface.

The engine replaces manual SAP screen extraction with a repeatable, parameterised workflow that ensures:

- consistent data definitions  
- shared business rules  
- synchronised refresh across multiple sheets  
- zero SQL or SAP knowledge required for end users  

---

## Key Features

- 🔌 **Direct ADO connection** to SAP SQL Server  
- 🎯 **Parameterised SQL queries** with shared filters  
- 🔄 **Single‑run refresh** of three interdependent operational views  
- 📐 **Dynamic resizing & formula propagation**  
- 🧩 **User‑controlled execution scope** via Excel inputs  
- 👥 Designed for **non‑technical operational teams**

---

## Architecture

### 1. User Input Layer
Users specify an **item‑code range** directly in Excel.

### 2. Execution Layer (VBA + ADO)
- Reads parameters from Excel  
- Builds parameterised SQL queries  
- Executes queries sequentially against SAP SQL Server  

### 3. Population Layer
- **PL&STR** — multi‑warehouse availability  
- **by Item** — PO quantities, ETD/ETA, shipment status  
- **by Container** — inbound shipment aggregation  

### 4. Governance Layer
- Clears previous records  
- Inserts new data  
- Resizes formulas to match row count  
- Ensures all views refresh together or not at all  

---

## Usage

1. Open the Excel file.  
2. Enter the **From / To item‑code range** in the green cells.  
3. Press **`Ctrl + Shift + Z`** or click the **Pivot** button.  
4. All operational views refresh in a **single governed run**.

---

## Repository Contents

```
/Excel_VBA.vba        # Anonymised full VBA implementation
/Excel_Sample.xlsx    # Sample Excel structure and UI
```

---

## Notes

This repository contains anonymised code and sample structures for **knowledge‑sharing purposes**.  
Production versions used in operational environments include additional governance and environment‑specific configuration.
