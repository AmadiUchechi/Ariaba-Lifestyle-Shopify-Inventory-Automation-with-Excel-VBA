# Ariaba-Lifestyle-Shopify-Inventory-Automation-with-Excel-VBA
Shopify Inventory Automation with Excel VBA

#  1. Project Overview

This project is an Excel VBA automation solution developed for Ariaba Lifestyle in August 2024. The goal was to simplify and automate the process of preparing, entering, updating, and managing product inventory data for the company's Shopify store.

Instead of manually entering and modifying large amounts of product information, I developed an Excel-based inventory management model using Visual Basic for Applications (VBA). The solution provides a structured product-entry form and automated functions for saving, updating, and deleting product records.

The project demonstrates how Excel, VBA programming, and process automation can be used to improve the efficiency and accuracy of e-commerce inventory management.

#  2. Business Problem

Ariaba Lifestyle needed a more efficient way to manage product information before uploading inventory to Shopify.

Managing Shopify product data manually can involve numerous fields, including:

-  Product titles
-  Product descriptions
-  SKUs
-  Product categories
-  Vendors
-  Tags
-  Inventory quantities
-  Prices
-  Compare-at prices
-  Product variants
-  Barcodes
-  Product images
-  Shipping information
-  Tax information
-  SEO information
-  Google Shopping information
-  Product status

Entering and modifying this information manually can be time-consuming and increase the possibility of errors.

The objective was therefore to create an Excel automation model that simplified product data entry and inventory updates while maintaining a structure compatible with Shopify product data.

#  3. Solution

I designed an Excel VBA-based product management system that allows product information to be entered through a user-friendly worksheet rather than directly manipulating the underlying product dataset.

The workbook contains a structured product form with fields corresponding to Shopify product attributes.

The VBA macros automate the movement and management of information between the input form and the underlying product data sheets.

The system includes three primary actions:

- SAVE — Adds product information entered into the form to the product dataset.

- UPDATE — Searches for an existing product using its unique Handle and modifies the corresponding product information.

- DELETE — Searches for a product by its Handle and removes the corresponding record from the product dataset.

#  4. Key Features
**Automated Product Entry**

The model provides a structured Excel form where users can enter product information without manually navigating through a large Shopify inventory spreadsheet.

The VBA program transfers the entered information to the appropriate fields in the product dataset.

**Inventory Updates**

Existing product information can be modified directly from the form.

The VBA program searches the product dataset using the product Handle as the identifier.

For example, the update functionality uses Excel VBA's .Find() method to locate the appropriate product record before modifying the relevant fields.


     
     Set rowToModify = wsData.Range("A:A").Find( _
     What:=Handle, _
     LookIn:=xlValues, _
     LookAt:=xlWhole)

Once the Handle is located, the macro updates the corresponding product information.

**Product Deletion**3

The model also allows users to delete an existing product record.

The VBA macro searches column A of the product dataset for the specified Handle and deletes the corresponding row.


    Set foundCell = wsData.Columns("A").Find( _
    What:=Handle, _
    LookIn:=xlValues, _
    LookAt:=xlWhole)

If Not foundCell Is Nothing Then
    
    foundCell.EntireRow.Delete
End If

Shopify-Compatible Product Fields

The form was designed around Shopify product data fields, including:

- Handle
- Title
- Body (HTML)
- Vendor
- Product Category
- Type
- Tags
- Published Status
- Published Scope
- Product Options
- Variant SKU
- Variant Inventory Quantity
- Variant Inventory Policy
- Variant Fulfillment Service
- Variant Price
- Variant Compare-at Price
- Variant Barcode
- Image Source
- Image Position
- Image Alt Text
- Gift Card
- Collections
- SEO Title
- SEO Description
- Google Shopping fields
- Cost per Item
- International Pricing
- Product Status

This structure makes it easier to organize product information for subsequent Shopify inventory operations.

#  5. Workbook Structure

The Excel solution uses multiple worksheets to separate the user interface from the underlying product data.

- Form

The Form worksheet acts as the main user interface.

Users enter product information into designated cells and use the command buttons to perform inventory operations.

The interface includes:

**SAVE** | **UPDATE** | **DELETE**

This approach makes the workbook easier to use for individuals who may not have experience working with VBA or complex Excel datasets.


- Products

The Products worksheet stores the structured product information.

The VBA macros communicate between the Form worksheet and the Products worksheet to locate and manipulate product records.

#  6. Raw Data

A Raw worksheet is also used within the automation process to organize data for further processing/export.

VBA Functions

The project includes VBA procedures responsible for the major inventory-management operations.


 **Save Product**
 |  ------  |
 
The Save macro:

- Collects information entered into the product form.
- Transfers the values into the appropriate product-data structure.
- Prepares the product information for inventory management/export.
- Reduces repetitive manual copying and pasting.
  

  **Update Product**
  |  ------  |
  
The ModifyDataByID() procedure:

- Reads the product Handle from the input form.
- Searches the Products worksheet for the Handle.
- Identifies the corresponding product row.
- Reads the updated values from the form.
- Replaces the appropriate values in the existing product record.
- Displays a confirmation message when the update is completed.
  

  **Delete Product**
  |  ------  |
  
The ClearDataByID() procedure:

- Reads the product Handle from the form.
- Searches the Products worksheet.
- Locates the matching product.
- Deletes the corresponding product row.
- Displays a confirmation message.
- Clears the Handle field after the operation.

#  Technologies Used


|     Technology     |     Purpose     |
|     ---          |    ---          |
|  Microsoft Excel     |     Inventory management interface     |
|  VBA	             |     Automation and business logic      |
|  Excel Macros	   |     Automated inventory operations     |
|  Shopify Product Structure	  |   Organization of e-commerce product information     |
|  Data Validation	   |     Controlled product data entry      |
|  Excel Forms/Buttons |     User interaction with VBA macros   |

#  7. Skills Demonstrated

This project demonstrates practical experience in:

- Excel VBA Programming
- Excel Automation
- Data Management
- Inventory Management
- E-commerce Data Management
- Shopify Product Data Preparation
- Process Automation
- Data Validation
- Conditional Logic
- VBA Loops
- Range Manipulation
- Excel .Find() Method
- CRUD-style Data Operations
- Business Process Improvement
- User-Friendly Excel Interface Design

#  Automation Workflow

The overall workflow of the application can be summarized as:

               ARIABA PRODUCT FORM
                       |
                       v
             Enter Product Details
                       |
          +------------+------------+
          |            |            |
          v            v            v
        SAVE         UPDATE       DELETE
          |            |            |
          v            v            v
     Add Product   Find Handle   Find Handle
          |            |            |
          v            v            v
     Store Data    Modify Data   Delete Record
          |            |            |
          +------------+------------+
                       |
                       v
               PRODUCT DATASET
                       |
                       v
            Shopify Data Preparation


#  8. Business Impact

The automation model was designed to reduce the amount of repetitive work involved in managing Ariaba Lifestyle's Shopify product information.

The solution helped provide:

- Faster product data entry
- Easier inventory updates
- Centralized product information
- Reduced repetitive copying and pasting
- More consistent product-data formatting
- Easier modification of existing inventory
- Reduced risk of manual data-entry errors
- A simpler interface for managing large numbers of Shopify product fields

Most importantly, the project transformed a repetitive inventory-management process into a button-driven Excel workflow

#  9. Project Screenshots
**Ariaba Product Management Form**

The Excel interface provides a centralized location for entering and managing Shopify product information.
<img width="1811" height="599" alt="Ariaba Product Management Form" src="https://github.com/user-attachments/assets/34c537ae-8c98-449f-9c40-90531e4119b2" />


**VBA Update Automation**

The VBA update procedure searches for a product using its Handle and modifies the corresponding inventory record.

<img width="1691" height="999" alt="VBA Update Automation" src="https://github.com/user-attachments/assets/175ceecb-aa72-4419-b291-a0e1a3b3974d" />


**VBA Delete Automation**

The delete procedure searches for the selected Handle and automatically removes the associated product record.

<img width="1753" height="995" alt="VBA Delete Automation" src="https://github.com/user-attachments/assets/752883ac-4895-451d-9b87-6f8181bf4c99" />



# 10. Challenges and Learning

One of the main challenges of this project was mapping a large number of Shopify product fields to a simple Excel input interface while ensuring that information was written to the correct locations.

The project strengthened my understanding of:

- Working with Excel ranges programmatically
- Searching datasets using VBA
- Automating repetitive business processes
- Updating records based on unique identifiers
- Creating reusable VBA procedures
- Designing Excel tools for non-technical users
- Structuring data for e-commerce applications

It also demonstrated how Excel can be transformed from a traditional spreadsheet into a lightweight business application using VBA.

#  11. Future Improvements

Possible improvements to the project include:

- Adding stronger input validation
- Automatically detecting duplicate product Handles
- Creating automated Shopify-ready CSV exports
- Adding error logging and exception handling
- Creating inventory-level alerts
- Adding product search functionality
- Developing an inventory dashboard
- Adding automated backups before updates or deletions
- Integrating directly with the Shopify API
- Migrating the solution to a database-backed application for larger inventories

#  Author

Uchechi Amadi

Data Analytics | Excel Automation | VBA | SQL | Business Process Automation

#  Project Date

August 2024

#  Disclaimer

This repository is intended to demonstrate the technical implementation and skills used in the project. Client-sensitive information, proprietary product data, credentials, and confidential business information have been excluded.

            
