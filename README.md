# Maritime Customs EDI Gateway: Geneva-to-Local Integration

This repository preserves the legacy implementation of a **Customs Declaration & Cargo Management System** developed for **Mediterranean Shipping Company (MSC)**. The system served as a critical middleware between MSC's Global Headquarters in **Geneva** and National Customs Authorities.

---

### 🏛️ Integration Workflow & Business Logic

The system was architected to automate the ingestion of international maritime data and its subsequent transformation into legally compliant customs declarations:

1.  **Data Ingestion:** Automated reception of **EDIFACT** files from Geneva containing global vessel manifests and cargo distribution data.
2.  **Code Equivalency Engine (Core Logic):** A specialized mapping layer designed to resolve discrepancies between Geneva's international coding standards and local **Customs Agency** requirements. 
    *   *Example:* Mapping global port codes, container types, and commodity categories to specific local regulatory equivalents.
3.  **Regulatory Output (CUSCAR/CUSRESP):** 
    *   **CUSCAR (Customs Cargo Report):** Generated the final electronic manifest for customs submission.
    *   **CUSRESP (Customs Response):** Parsed official feedback from customs to confirm acceptance or flag errors.

---

### 📂 Data Structures & EDI Components

The system processed complex pipe-delimited and fixed-width formats to ensure full traceability:

*   **`ENVCTROL.txt` (Audit & Control):** A transmission header that calculated **Hash Totals** and record counts to ensure 100% data parity between Geneva and the local transmission.
*   **`MANHDR.txt` (Vessel Header):** Handled vessel and voyage identification logic (e.g., *MSC ACAPULCO*).
*   **`MANDET.txt` (Cargo Detail):** Itemized container details, weights, and logistics routing.
*   **`MANCON.txt` (Container Logic):** Specific mapping for container IDs (MSCU, GLDU), seals, and types.

---

### 🛠️ Technical Stack (Legacy)

*   **Language:** Visual Basic 6 (VB6).
*   **Format:** EDIFACT / Flat-file (Pipe-delimited and Fixed-width).
*   **Architecture:** N-Tier logic focused on high-reliability string manipulation and transactional integrity.

---

### 🚀 Key Engineering Challenges

*   **Complex Mapping:** Designing a flexible equivalency matrix to handle evolving customs codes without breaking the global data feed from Geneva.
*   **Data Integrity:** Implementing automated checksums (within `ENVCTROL`) to prevent financial or logistical discrepancies during transformation.
*   **Legacy Interoperability:** Bridging the gap between the shipping line's internal mainframe data and the government's electronic gateways.

---

> [!NOTE]
> All sensitive commercial identifiers and proprietary Geneva connection protocols have been removed. This source code is intended to showcase **Architectural Logic, Data Mapping Strategy, and Technical Problem-Solving**.

---
*Developed by Carlos Reyes - Senior .NET Developer & Business Systems Analyst*
