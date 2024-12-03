<h1 align="center">Null Summary Report Automation</h1>

  <h1 align="center">
  <img src="https://media0.giphy.com/media/xvBv5pU4djudjF0ri8/giphy.gif?cid=6c09b952ifm9y3pwlgnlx1le5wa9c2a7o35d52rtuw7ces39&ep=v1_internal_gif_by_id&rid=giphy.gif&ct=s" alt="Excel Icon" width="50" style="vertical-align">
</h1>

<h3 align="center">
This automation drastically simplify's the Null Location LPNs Report processing by formulating the data classification and reporting procedure using a VBA script in Microsoft Excel. The workflow efficiently manages LPN data by cleaning, filtering, and then transferring the data between workbooks, ensuring an accurate, streamlined and easy to read report.
</h3>

  
## **Project Overview**
### **Stage 1: Main Stage**  
Can be run on the **Null Locations LPNs Workbook** individually, producing the desired report with minimal effort.  
- **Note:** Rows are not deleted; only `SHIFT` and `DEPT` are formulated. Adjust target cells accordingly.

### **Stage 2: Data Removal Stage**  
Removes **Rising Son records (582)** from the report.

### **Stage 3: Data Transfer Stage**  
Transfers data to a developed workbook and adds a styling component for enhanced presentation.

---

## **Features**
### **Data Cleanup**
- Deletes rows based on specific criteria (e.g., warehouse codes, date ranges).  
- Removes unnecessary columns for a streamlined dataset.

### **Data Transformation**
- Dynamically adjusts worksheet names and structures.  
- Filters data based on inventory aging policies (e.g., records younger than 3 days).

### **Data Transfer**
- Transfers filtered data between source and destination workbooks.  
- Ensures existing headers in the destination workbook are not overwritten.

### **Workflow Automation**
- Combines all macros into a single workflow for sequential processing.

---

## **Prerequisites**
- **Microsoft Excel** with VBA macros enabled.  
- Familiarity with enabling macros in Excel (Developer tab must be enabled).  
- Source workbook naming convention starts with *"Null Location LPNs"* or similar.

---

## **Setup**
### **Enable Macros**  
1. Open Excel.  
2. Navigate to **File > Options > Trust Center > Trust Center Settings > Macro Settings**.  
3. Enable macros and allow VBA access.

### **Load VBA Macros**  
1. Open the **Visual Basic for Applications** editor (`Alt + F11`).  
2. Import the provided VBA scripts into the desired workbook.

### **Save as Macro-Enabled Workbook**  
- Save the workbook as a `.xlsm` file to preserve macros.

---

## **Macro Overview**
### **1. NullSummaryReport_Stage1**  
- Prepares the dataset by performing initial cleanup and transformations.

### **2. NullSummaryReport_Stage2**  
- Filters and deletes data based on warehouse codes and inventory aging policies.  
- Removes specific columns and renames the worksheet.

### **3. NullSummaryReport_Stage3**  
- Transfers data from the source workbook (*My Null Location LPNs*) to the destination workbook (*Sandbox*), starting from row 4.

### **4. RunAllStages**  
- Executes all stages sequentially for a complete workflow.

---

## **Execution Instructions**
1. Ensure the source workbook is open.  
2. Run the desired macro(s) via the VBA editor or assign a button in the workbook for ease of use.  
3. Review the output for accuracy and adjust criteria as needed.

---

## **Troubleshooting**
### **Workbook Not Found**
- Verify that the source workbook is open before running macros.  
- Ensure the naming convention matches the code (e.g., *"Null Location LPNs"*).

### **Incorrect Data Deletion**
- Review the filter logic in **NullSummaryReport_Stage2** and adjust criteria as needed.

### **Macro Errors**
- Ensure macros are enabled and saved in a macro-enabled workbook (`.xlsm`).  
- Confirm worksheet names match the references in the macros.
