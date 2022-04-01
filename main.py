from openpyxl import Workbook
import codecs
import pandas as pd

wb = Workbook()
ws = wb.active


with codecs.open('csv\\tit_ap_adiantamento.csv',encoding='utf-8') as csv_file:

        row = []       
       
        row.append("")#*Invoice ID
        row.append("")#*Business Unit
        row.append("")#*Source
        row.append("")#*Invoice Number
        row.append("")# *Invoice Amount
        row.append("")# *Invoice Date
        row.append("")# **Supplier Name
        row.append("")# **Supplier Number
        row.append("")# *Supplier Site
        row.append("")# Invoice Currency 
        row.append("")# Payment Currency
        row.append("")# Description
        row.append("")# Import Set
        row.append("")# *Invoice Type
        row.append("")# Legal Entity
        row.append("")# Customer Tax Registration Number
        row.append("")# Customer Registration Code
        row.append("")# First-Party Tax Registration Number
        row.append("")# Supplier Tax Registration Number
        row.append("")# *Payment Terms
        row.append("")# Terms Date
        row.append("")# Goods Received Date
        row.append("")# Invoice Received Date
        row.append("")# Accounting Date
        row.append("")# Payment Method
        row.append("")# Pay Group
        row.append("")# Pay Alone
        row.append("")# Discountable Amount
        row.append("")# Prepayment Number
        row.append("")# Prepayment Line Number
        row.append("")# Prepayment Application Amount
        row.append("")# Prepayment Accounting Date
        row.append("")# Invoice Includes Prepayment
        row.append("")# Conversion Rate Type
        row.append("")# Conversion Date
        row.append("")# Conversion Rate
        row.append("")# Liability Combination
        row.append("")# Document Category Code
        row.append("")# Voucher Number
        row.append("")# Requester First Name
        row.append("")# Requester Last Name
        row.append("")# Requester Employee Number
        row.append("")# Delivery Channel Code
        row.append("")# Bank Charge Bearer
        row.append("")# Remit-to Supplier
        row.append("")# Remit-to Supplier Number
        row.append("")# Remit-to Address Name
        row.append("")# Payment Priority
        row.append("")# Settlement Priority
        row.append("")# Unique Remittance Identifier
        row.append("")# Unique Remittance Identifier Check Digit
        row.append("")# Payment Reason Code
        row.append("")# Payment Reason Comments
        row.append("")# Remittance Message 1
        row.append("")# Remittance Message 2
        row.append("")# Remittance Message 3
        row.append("")# Withholding Tax Group
        row.append("")# Ship-to Location
        row.append("")# Taxation Country
        row.append("")# Document Sub Type
        row.append("")# Tax Invoice Internal Sequence Number
        row.append("")# Supplier Tax Invoice Number
        row.append("")# Tax Invoice Recording Date
        row.append("")# Supplier Tax Invoice Date
        row.append("")# Supplier Tax Invoice Conversion Rate
        row.append("")# Port Of Entry Code
        row.append("")# Correction Year
        row.append("")# Correction Period
        row.append("")# Import Document Number
        row.append("")# Import Document Date
        row.append("")# Tax Control Amount
        row.append("")# Calculate Tax During Import
        row.append("")# Add Tax To Invoice Amount
        row.append("")# Attribute Category
        row.append("")# Attribute 1
        row.append("")# Attribute 2
        row.append("")# Attribute 3
        row.append("")# Attribute 4
        row.append("")# Attribute 5
        row.append("")# Attribute 6
        row.append("")# Attribute 7
        row.append("")# Attribute 8
        row.append("")# Attribute 9
        row.append("")# Attribute 10
        row.append("")# Attribute 11
        row.append("")# Attribute 12
        row.append("")# Attribute 13
        row.append("")# Attribute 14
        row.append("")# Attribute 15
        row.append("")# Attribute Number 1
        row.append("")# Attribute Number 2
        row.append("")# Attribute Number 3
        row.append("")# Attribute Number 4
        row.append("")# Attribute Number 5
        row.append("")# Attribute Date 1
        row.append("")# Attribute Date 2
        row.append("")# Attribute Date 3
        row.append("")# Attribute Date 4
        row.append("")# Attribute Date 5
        row.append("")# Global Attribute Category
        row.append("")# Global Attribute 1

        ws.append(row)
        
           
wb.save("Plan_Contas_AP_ADIANTAMENTO_20220331_GOLIVE.xlsx")

    