import tkinter as tk
import tkinter.messagebox as messagebox
from PyPDF2 import PdfReader
import re
import pandas as pd
import os


def convert_folder_to_csv():
    try:
        # Get the folder path and file name from the entry widgets
        lst = []
        folder_path = folder_path_entry.get()
        out_file = file_name_entry.get()
        for file_name in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file_name)
            if os.path.isfile(file_path) and file_path.lower().endswith('.pdf'):
                lst.append(os.path.abspath(file_path))
    
        csv_data = []
        for ele in lst:
            reader = PdfReader(ele)
            text = ""
            for i in range(len(reader.pages)):
                page = reader.pages[i]
                text += page.extract_text().replace("Page 1 of {page}".format(page=len(reader.pages)), "").replace(
                    "Page 2 of {page}".format(page=len(reader.pages)), "").replace("\n", " ")
        
            try:
                data = {"CSB Number": re.search(r'(?<=CSB Number:\s).+?(?=Filling Date)', text).group(0).strip(),
                        "Filling Date": re.search(r'(?<=Filling Date:)+\s*\d{1,2}\/\d{1,2}\/\d{4}', text).group(0).strip(),
                        "Courier Registration Number": re.search(r'(?<=Courier Registration Num- ber:).+?(?=Courier Name:)', text).group(0).strip(),
                        "Courier name": re.search(r'(?<=Courier Name:).+?(?=Address)', text).group(0).strip(),
                        "Address 1": re.search(r'(?<=Address 1:\s).+?(?=\sAddress)', text).group(0).strip(),
                        "Address 2": re.search(r'(?<=Address 2:\s).+?(?=City)', text).group(0).strip(),
                        "City": re.search(r'(?<=City:).+?(?=Postal)', text).group(0).strip(),
                        "Postal zip code": re.search(r'(?<=Zip Code:\s).+?(?=State)', text).group(0).strip(),
                        "State": re.search(r'(?<=State:).+?(?=Custom)', text).group(0).strip(),
                        "Custom Station Name": re.search(r'(?<=Custom Station Name).+?(?=AIRLINE)', text).group(0).strip(),
                        "Airlines": re.search(r'(?<=Airlines:).+?(?=Flight)', text).group(0).strip(),
                        "Flight Number": re.search(r'(?<=Flight Number:).+?(?=Port)', text).group(0).strip(),
                        "Port of Loading": re.search(r'(?<=Port of Loading:).+?(?=Date)', text).group(0).strip(),
                        "HAWB Number": re.search(r'(?<=HAWB Number:).+?(?=Number)', text).group(0).strip(),
                        "Number of Packages": re.search(r'(?<=ULD:).+?(?=Declared)', text).group(0).strip(),
                        "Declared Weight(in Kgs)": re.search(r'(?<=Declared Weight\(in Kgs\):).+?(?=Airport)', text).group(0).strip(),
                        "Airport of Destination": re.search(r'(?<=Airport of Destination:).+?(?=Import)', text).group(0).strip(),
                        "Import Export Code": re.search(r'(?<=Import Export Code \(IEC\):).+?(?=IEC)', text).group(0).strip(),
                        "IEC Branch Code": (((re.search(r'(?<=IEC Branch Code:).+?(?=Invoice Term)', text).group(0).strip()) if (re.search(r'(?<=IEC Branch Code:).+?(?=Invoice Term)', text)) else " ")),
                        "Invoice Term": (((re.search(r'(?<=Invoice Term:).+?(?=MHBS)', text).group(0).strip()) if (re.search(r'(?<=Invoice Term:).+?(?=MHBS)', text)) else " ")),
                        "MHBS No": (((re.search(r'(?<=MHBS No:).+?(?=Export Using e-Commerce)', text).group(0).strip()) if (re.search(r'(?<=MHBS No:).+?(?=Export Using e-Commerce)', text)) else " ")),
                        "Export Using e-Commerce": (((re.search(r'(?<=Export Using e-Commerce:).+?(?=Under ME)', text).group(0).strip()) if (re.search(r'(?<=Export Using e-Commerce:).+?(?=Under ME)', text)) else " ")),
                        "Under MEIS Scheme": (((re.search(r'(?<=Under MEIS Scheme:).+?(?=if Export using)', text).group(0).strip()) if (re.search(r'(?<=Under MEIS Scheme:).+?(?=if Export using)', text)) else " ")),
                        "if Export using E-Commerce is 'Yes' and the export consignment contains jewellery falling under CTH 7117 or 7113:": (((re.search(r'(?<= 7117 or7113:).+?(?=(i) Name of)', text).group(0).strip()) if (re.search(r'(?<= 7117 or7113:).+?(?=(i) Name of)', text)) else " ")),
                        "Name of E-commerce Operator or Website ": (((re.search(r'(?<=Website :).+?(?=\(iii\)\ Order No)', text).group(0).strip()) if (re.search(r'(?<=Website :).+?(?=\(iii\)\ Order No)', text)) else " ")),
                        "Order No ": (((re.search(r'(?<=Order No :).+?(?=\(ii\)\ Payment)', text).group(0).strip()) if (re.search(r'(?<=Order No :).+?(?=\(ii\)\ Payment)', text)) else " ")),
                        "Unique transaction ID": (((re.search(r'(?<=tion ID :).+?(?=\(iv\)\ Order)', text).group(0).strip()) if (re.search(r'(?<=tion ID :).+?(?=\(iv\)\ Order)', text)) else " ")),
                        "Order Date": (((re.search(r'(?<=Order Date :).+?(?=AD Code:)', text).group(0).strip()) if (re.search(r'(?<=Order Date :).+?(?=AD Code:)', text)) else " ")),
                        "AD Code": (((re.search(r'(?<=AD Code:).+?(?=Account No:)', text).group(0).strip()) if (re.search(r'(?<=AD Code:).+?(?=Account No:)', text)) else " ")),
                        "Account No": (((re.search(r'(?<=Account No:).+?(?=Government\/)', text).group(0).strip()) if (re.search(r'(?<=Account No:).+?(?=Government\/)', text)) else " ")),
                        "Government / Non-Government": (((re.search(r'(?<=Government\/ Non-Government:).+?(?=NFEI)', text).group(0).strip()) if (re.search(r'(?<=Government\/ Non-Government:).+?(?=NFEI)', text)) else " ")),
                        "Status": (((re.search(r'(?<=Status:).+?(?=LEO DATE:)', text).group(0).strip()) if (re.search(r'(?<=Status:).+?(?=LEO DATE:)', text)) else " ")),
                        "LEO DATE": re.search(r'(?<=LEO DATE:)+\s*\d{1,2}\/\d{1,2}\/\d{2}', text).group(0).strip(),
                        "FOB Value (In INR)": (((re.search(r'(?<=FOB Value \(In INR\):).+?(?=FOB Value)', text).group(0).strip()) if (re.search(r'(?<=FOB Value \(In INR\):).+?(?=FOB Value)', text)) else " ")),
                        "FOB Value (In Foreign Currency)": (((re.search(r'(?<=FOB Value \(In Foreign Cur- rency\):).+?(?=FOB Exchange Rate)', text).group(0).strip()) if (re.search(r'(?<=FOB Value \(In Foreign Cur- rency\):).+?(?=FOB Exchange Rate)', text)) else " ")),
                        "FOB Exchange Rate (In Foreign Currency)": (((re.search(r'(?<=eign Currency\):).+?(?=FOB Currency)', text).group(0).strip()) if (re.search(r'(?<=eign Currency\):).+?(?=FOB Currency)', text)) else " ")),
                        "FOB Currency (In Foreign Currency)": (((re.search(r'(?<=FOB Currency \(In Foreign Currency\):).+?(?=CEM DETAILS)', text).group(0).strip()) if (re.search(r'(?<=FOB Currency \(In Foreign Currency\):).+?(?=CEM DETAILS)', text)) else " ")),
                        "EGM Number": (((re.search(r'(?<=EGM Number:).+?(?=EGM Date)', text).group(0).strip()) if (re.search(r'(?<=EGM Number:).+?(?=EGM Date)', text)) else " ")),
                        "EGM Number": (((re.search(r'(?<=EGM Number:).+?(?=EGM Date)', text).group(0).strip()) if (re.search(r'(?<=EGM Number:).+?(?=EGM Date)', text)) else " ")),
                        "EGM Date": (((re.search(r'(?<=EGM Date:).+?(?=CONSIGNOR AND CONSIGNEE)', text).group(0).strip()) if (re.search(r'(?<=EGM Date:).+?(?=CONSIGNOR AND CONSIGNEE)', text)) else " ")),
                        "Name of the Consignor": (((re.search(r'(?<=Name of the Consignor:).+?(?=Address of the Consignor)', text).group(0).strip()) if (re.search(r'(?<=Name of the Consignor:).+?(?=Address of the Consignor)', text)) else " ")),
                        "Address of the Consignor": (((re.search(r'(?<=Address of the Consignor:).+?(?=Name of the Consignee)', text).group(0).strip()) if (re.search(r'(?<=Address of the Consignor:).+?(?=Name of the Consignee)', text)) else " ")),
                        "Name of the Consignee": (((re.search(r'(?<=Name of the Consignee:).+?(?=Address of the Consignee)', text).group(0).strip()) if (re.search(r'(?<=Name of the Consignee:).+?(?=Address of the Consignee)', text)) else " ")),
                        "Address of the Consignee": (((re.search(r'(?<=Address of the Consignee:).+?(?=GSTIN DETAILS)', text).group(0).strip()) if (re.search(r'(?<=Address of the Consignee:).+?(?=GSTIN DETAILS)', text)) else " ")),
                        "KYC Document": (((re.search(r'(?<=KYC Document:).+?(?=KYC ID)', text).group(0).strip()) if (re.search(r'(?<=KYC Document:).+?(?=KYC ID)', text)) else " ")),
                        "KYC ID": (((re.search(r'(?<=KYC ID:).+?(?=State Code)', text).group(0).strip()) if (re.search(r'(?<=KYC Document:).+?(?=State Code)', text)) else " ")),
                        "State Code": (((re.search(r'(?<=State Code:).+?(?=GSTIN Uploaded)', text).group(0).strip()) if (re.search(r'(?<=KYC Document:).+?(?=GSTIN Uploaded)', text)) else " ")),
                        "GSTIN Uploaded To ICEGATE Server": (((re.search(r'(?<=GSTIN Uploaded To ICEGATE Server:).+?(?=INVOICE DETAILS)', text).group(0).strip()) if (re.search(r'(?<=GSTIN Uploaded To ICEGATE Server:).+?(?=INVOICE DETAILS)', text)) else " ")),
                        "CRN Number": (((re.search(r'CRN MHBS Number:\s*([A-Za-z0-9]+)', text).group(1).strip()) if (re.search(r'CRN MHBS Number:\s*([A-Za-z0-9]+)', text)) else " ")),
                        "CRN Number": (((re.search(r'CRN MHBS Number:\s([A-Za-z0-9]+)', text).group(1).strip()) if (re.search(r'CRN MHBS Number:\s*([A-Za-z0-9]+)', text)) else " ")),
                        "CRN MHBS Number": (((re.search(r'CRN MHBS Number:\s[A-Za-z0-9]+\s([a-zA-Z0-9]+).+?(?=DECLARATION)', text).group(1).strip()) if (re.search(r'CRN MHBS Number:\s[A-Za-z0-9]+\s([a-zA-Z0-9]+).+?(?=DECLARATION)', text)) else " ")),

                        }
                length = len(re.findall("ITEM DETAILS", text))
                for i in range(length):
                    item_data = {
                        "Invoice Number "+str(i+1): ((re.findall(r'(?<=Invoice Value \(in INR\):).+?(?=\d{1,2}\/\d{1,2}\/\d{4})', text)[i].strip() if (re.search(r'(?<=Invoice Value \(in INR\):).+?(?=\d{1,2}\/\d{1,2}\/\d{4})', text)) else " ")),
                        "Invoice Date "+str(i+1): ((re.findall(r'Invoice Value \(in INR\):\s*\d*\s*(\d{1,2}\/\d{1,2}\/\d{4})', text)[i].strip() if (re.search(r'Invoice Value \(in INR\):\s*\d*\s*(\d{1,2}\/\d{1,2}\/\d{4})', text)) else " ")),
                        "Invoice value "+str(i+1): ((re.findall(r'Invoice Value \(in INR\):\s*\d*\s*\d{1,2}\/\d{1,2}\/\d{4}\s*(\d*)', text)[i].strip() if (re.search(r'Invoice Value \(in INR\):\s*\d*\s*\d{1,2}\/\d{1,2}\/\d{4}\s*(\d*)', text)) else " ")),
                        "CTSH "+str(i+1): ((re.findall(r'(?<=CTSH:).+?(?=\(ii\) SKU)', text)[i].strip() if (re.search(r'(?<=CTSH:).+?(?=\(ii\) SKU)', text)) else " ")),
                        "SKU NO. "+str(i+1): ((re.findall(r'(?<=SKU NO :).+?(?=\(iii\) Type)', text)[i].strip() if (re.search(r'(?<=SKU NO :).+?(?=\(iii\) Type)', text)) else " ")),
                        "Type of Jewellery "+str(i+1): ((re.findall(r'(?<=Type of Jewellery :).+?(?=Goods Description)', text)[i].strip() if (re.search(r'(?<=Type of Jewellery :).+?(?=Goods Description)', text)) else " ")),
                        "Goods Description "+str(i+1): ((re.findall(r'(?<=Goods Description:).+?(?=Quantity)', text)[i].strip() if (re.search(r'(?<=Goods Description:).+?(?=Quantity)', text)) else " ")),
                        "Quantity "+str(i+1): ((re.findall(r'(?<=Quantity:).+?(?=Unit Of Measure)', text)[i].strip() if (re.search(r'(?<=Quantity:).+?(?=Unit Of Measure)', text)) else " ")),
                        "Unit Of Measure "+str(i+1): ((re.findall(r'(?<=Unit Of Measure:).+?(?=Unit Price)', text)[i].strip() if (re.search(r'(?<=Unit Of Measure:).+?(?=Unit Price)', text)) else " ")),
                        "Unit Price "+str(i+1): ((re.findall(r'(?<=Unit Price:).+?(?=Total Item Value)', text)[i].strip() if (re.search(r'(?<=Unit Price:).+?(?=Total Item Value)', text)) else " ")),
                        "Total Item Value "+str(i+1): ((re.findall(r'(?<=Total Item Value:).+?(?=Unit Price Currency)', text)[i].strip() if (re.search(r'(?<=Total Item Value:).+?(?=Unit Price Currency)', text)) else " ")),
                        "Unit Price Currency "+str(i+1): ((re.findall(r'(?<=Unit Price Currency:).+?(?=Exchange Rate)', text)[i].strip() if (re.search(r'(?<=Unit Price Currency:).+?(?=Exchange Rate)', text)) else " ")),
                        "Exchange Rate "+str(i+1): ((re.findall(r'(?<=Exchange Rate:).+?(?=Total Item Value)', text)[i].strip() if (re.search(r'(?<=Exchange Rate:).+?(?=Total Item Value)', text)) else " ")),
                        "Total Item Value (In INR) "+str(i+1): ((re.findall(r'(?<=Total Item Value \(In INR\):).+?(?=Total Taxable Value)', text)[i].strip() if (re.search(r'(?<=Total Item Value \(In INR\):).+?(?=Total Taxable Value)', text)) else " ")),
                        "Taxable Value Currency "+str(i+1): ((re.findall(r'(?<=Taxable Value Currency:).+?(?=Total IGST Paid)', text)[i].strip() if (re.search(r'(?<=Taxable Value Currency:).+?(?=Total IGST Paid)', text)) else " ")),
                        "Total IGST Paid "+str(i+1): ((re.findall(r'(?<=Total IGST Paid:).+?(?=BOND OR UT)', text)[i].strip() if (re.search(r'(?<=Total IGST Paid:).+?(?=BOND OR UT)', text)) else " ")),
                        "BOND OR UT "+str(i+1): ((re.findall(r'(?<=BOND OR UT:).+?(?=Total CESS Paid)', text)[i].strip() if (re.search(r'(?<=BOND OR UT:).+?(?=Total CESS Paid)', text)) else " ")),
                        "Total CESS Paid "+str(i+1): ((re.findall(r'(?<=Total CESS Paid:).+?(?=Purity)', text)[i].strip() if (re.search(r'(?<=Total CESS Paid:).+?(?=Purity)', text)) else " ")),
                        "Purity "+str(i+1): ((re.findall(r'(?<=Purity :).+?(?=Whether studded)', text)[i].strip() if (re.search(r'(?<=Purity :).+?(?=Whether studded)', text)) else " ")),
                        "Whether studded or set with precious/semiprecious stones "+str(i+1): ((re.findall(r'(?<=semiprecious stones :).+?(?=Wt.)', text)[i].strip() if (re.search(r'(?<=semiprecious stones :).+?(?=Wt.)', text)) else " ")),
                        "If yes "+str(i+1): ((re.findall(r'(?<=If yes, :).+?(?=\(a\) Diamond)', text)[i].strip() if (re.search(r'(?<=If yes, :).+?(?=\(a\) Diamond)', text)) else " ")),
                        "Wt.(in gm) "+str(i+1): ((re.findall(r'(?<=Wt.\(in gm\):).+?(?=If yes)', text)[i].strip() if (re.search(r'(?<=Wt.\(in gm\):).+?(?=If yes)', text)) else " ")),
                        "Diamond Cut "+str(i+1): ((re.findall(r'(?<=Cut :).+?(?=\(b\) Color)', text)[i].strip() if (re.search(r'(?<=Cut :).+?(?=\(b\) Color)', text)) else " ")),
                        "Diamond Color "+str(i+1): ((re.findall(r'(?<=Color :).+?(?=\(c\) Clarity)', text)[i].strip() if (re.search(r'(?<=Color :).+?(?=\(c\) Clarity)', text)) else " ")),
                        "Diamond Clarity "+str(i+1): ((re.findall(r'(?<=Clarity :).+?(?=\(d\) Carat)', text)[i].strip() if (re.search(r'(?<=Clarity :).+?(?=\(d\) Carat)', text)) else " ")),
                        "Diamond No.of.Stones "+str(i+1): ((re.findall(r'(?<=\(e\) No.of.Stones :).+?(?=\(b\) If other precious)', text)[i].strip() if (re.search(r'(?<=\(e\) No.of.Stones :).+?(?=\(b\) If other precious)', text)) else " ")),
                        "Name of the stone "+str(i+1): ((re.findall(r'(?<=Name of the stone :).+?(?=\(b\) Whether Natural)', text)[i].strip() if (re.search(r'(?<=Name of the stone :).+?(?=\(b\) Whether Natural)', text)) else " ")),
                        "Whether Natural or Synthetic "+str(i+1): ((re.findall(r'(?<=thetic :).+?(?=\(c\) No. of Stones)', text)[i].strip() if (re.search(r'(?<=thetic :).+?(?=\(c\) No. of Stones)', text)) else " ")),
                        "No. of Stones "+str(i+1): ((re.findall(r'(?<=No. of Stones:).+?(?=\(d\) Country of origin)', text)[i].strip() if (re.search(r'(?<=No. of Stones:).+?(?=\(d\) Country of origin)', text)) else " ")),
                        "Country of origin "+str(i+1): ((re.findall(r'(?<=Country of origin :).+?(?=[Invoice Number]|[CRN DETAILS])', text)[i].strip() if (re.search(r'(?<=Country of origin :).+?(?=[Invoice Number]|[CRN DETAILS])', text)) else " ")),
                    }
                    data.update(item_data)
                csv_data.append(data)

            except Exception as e:
                pass
        file_path = os.path.join(folder_path, out_file+".csv")
        df = pd.DataFrame(csv_data)
        df.to_csv(file_path, index=False, header=True)
        messagebox.showinfo('Success', 'Conversion was successful!')

    except Exception as e:
        messagebox.showinfo('Error', str(e))


# Create a new Tkinter window
window = tk.Tk()

window.geometry('800x600')
# Set the title of the window
window.title('PDFs to CSV Converter')

# Create a label and entry widget for the folder path
folder_path_label = tk.Label(window, text='Folder path:')
folder_path_label.grid(row=0, column=0)
folder_path_entry = tk.Entry(window)
folder_path_entry.grid(row=0, column=1)

# Create a label and entry widget for the file name
file_name_label = tk.Label(window, text='Output File name:')
file_name_label.grid(row=1, column=0)
file_name_entry = tk.Entry(window)
file_name_entry.grid(row=1, column=1)

# Create a button that triggers the conversion when clicked
convert_button = tk.Button(window, text='Convert',
                           command=convert_folder_to_csv)
convert_button.grid(row=2, column=0, columnspan=2)

# Start the Tkinter event loop
window.mainloop()
