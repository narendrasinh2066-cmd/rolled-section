import streamlit as st
import pandas as pd
import openpyxl

st.title("Rolled Section - Automation UI")
st.write("Apni details select karein aur table me data bharein")

# 1. E4 aur E6 ke Dropdowns
# Yahan brackets [] ke andar aap apne baki drawing no aur sections add kar sakte hain
drawing_no = st.selectbox("Select Drawing No (Excel Cell: E4):", ["Drawing-1", "Drawing-2", "Drawing-3"]) 
section = st.selectbox("Select Section (Excel Cell: E6):", ["ISMB200", "ISA75X75X6", "ISMC100", "ISMB250"])

# 2. D10 aur E10 ke liye Excel jaisi Table banana
st.write("Enter Length (D) and QTY (E):")
# Starting me 5 khali rows dikhayenge (aap isko bada/ghata sakte hain)
df_empty = pd.DataFrame([{"Length": 0.0, "QTY": 0}] * 5)

# st.data_editor se aap mobile me hi directly values type kar payenge
edited_data = st.data_editor(df_empty, num_rows="dynamic", use_container_width=True)

# 3. Submit Button
if st.button("Run Python Logic"):
    try:
        # Excel file background me open karna
        wb = openpyxl.load_workbook('Rolled Section.xlsx')
        sheet = wb['Automation'] # Aapki sheet ka naam

        # Dropdown ki values ko E4 aur E6 me dalna
        sheet['E4'] = drawing_no
        sheet['E6'] = section

        # Pehle ka purana data clear karna (D10 se E100 tak) taaki purani lines na bachein
        for r in range(10, 100):
            sheet.cell(row=r, column=4).value = None # Column D (Length)
            sheet.cell(row=r, column=5).value = None # Column E (QTY)

        # Naya data table se padh kar D10 aur E10 se niche likhna shuru karna
        current_row = 10
        for index, row in edited_data.iterrows():
            length_val = row['Length']
            qty_val = row['QTY']
            
            # Agar length 0 se jyada daali hai, tabhi Excel me add karein
            if length_val > 0:
                sheet.cell(row=current_row, column=4).value = length_val 
                sheet.cell(row=current_row, column=5).value = qty_val    
                current_row += 1

        # File save karna
        new_filename = 'Updated_Rolled_Section.xlsx'
        wb.save(new_filename)
        st.success("Data Excel me add ho gaya hai!")
        st.write(f"Total Rows Added: {current_row - 10}")

        # ==========================================
        # YAHAN AAPKA MAIN PYTHON CODE AAYEGA 
        # Jo "Updated_Rolled_Section.xlsx" par kaam kare
        # ==========================================

        # Download button taaki aap check kar sakein ki data sahi jagah gaya ya nahi
        with open(new_filename, "rb") as file:
            st.download_button(label="Download Updated Excel File", data=file, file_name=new_filename)
            
    except Exception as e:
        st.error(f"Error aayi hai: {e}. Kripya check karein ki file me 'Automation' naam ki sheet mojood hai.")
