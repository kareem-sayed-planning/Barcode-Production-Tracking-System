# 📟 Barcode Production Tracking System (VBA Data Pipeline)

![Body Shop Scanner](Screenshot%202026-06-14%20184251.png)

![Final OK Scanner](Screenshot%202026-06-14%20190315.png)

## 📝 Overview
This project serves as the foundational data collection layer for the production management ecosystem. It is an enterprise-grade Excel VBA application deployed across multiple shop floors (Body, Paint, Assembly, Final) to capture real-time vehicle movement using barcode scanners. This system eliminates manual data entry, acting as the primary data feed for the Live Monitoring Dashboard.

## 🚀 Key Features & Business Logic
* **Real-Time Barcode Capture:** Validates 17-character VINs instantly to prevent operator errors.
* **Automated Shift Calculation:** Dynamically assigns records to Shift 1 or Shift 2 based on the actual scan time and daily working hour parameters.
* **Smart Data Consolidation:** Utilizes advanced VBA Dictionaries and Arrays to process thousands of records across multiple shared workbooks simultaneously without lag.
* **Duplicate Prevention:** Cross-references live inputs against historical databases to block duplicate entries in real-time.

## 🛠️ Tech Stack & Architecture
* **Language:** Excel VBA (Macros)
* **Architecture:** Decentralized UserForms (Scanners) feeding into a Centralized Master Database (`SSAR`).
* **Performance:** Optimized using in-memory Arrays and `Scripting.Dictionary` for high-speed read/write operations.

---
*This system represents the Data Collection phase (Phase 1) of the overall Manufacturing Execution process.*
