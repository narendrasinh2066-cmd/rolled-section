import streamlit as st
import pandas as pd
import openpyxl

st.title("Rolled Section - Automation UI")
st.write("Apni details select karein aur table me data bharein")

# 1. Exact Drawing Numbers from your screenshot
drawing_list = [
    "MG500-CVS-500PC-LC-0472", "MG500-CVC-500PG-LC-0529", "MG300-CVS-301MH-LC-0184", 
    "MGA00-CVS-A03PZ-LC-0001", "MG300-CVS-301UT-LC-0003", "MG600-CVS-602PC-LC-1912", 
    "MG600-CVS-602PC-LC-1872", "MG600-CVS-602AA-LC-0327", "MG300-CVS-301PP-LC-0027", 
    "MG300-CVS-301PP-LC-0025", "MG600-CVS-602AG-LC-1251", "MG600-CVS-602AH-LC-1627", 
    "MGA00-CVS-A03QJ-LC-0001", "MGA00-CVC-A06SE-LC-0020", "MGA00-CVC-A03QL-LC-0001", 
    "MG500-CVC-502PA-LC-0500", "MGA00-CVS-A03QN-LC-0002", "MGA00-CVS-A03PG-LC-0003", 
    "MGA00-CVC-A03PF-LC-0002", "MG500-CVS-501AE-LC-1070", "MG500-CVC-502AD-LC-0301", 
    "MG300-CVS-301MH-LC-0050", "MG400-CVS-401EB-LC-5801", "MG400-CVS-400CC-LC-0601", 
    "MG400-CVS-400CC-LC-0602", "MG500-CVS-500SA-LC-0668", "MG500-CVS-500AB-LC-1204", 
    "MG500-CVS-501AF-LC-0744", "MG500-CVS-500AB-LC-1204", "MGA00-CVS-A01AD-LC-0006", 
    "MGA00-CVC-002AA-LC-7004", "MGA00-CVC-002AA-LC-7005"
]

# 2. Exact Section Profiles from your screenshot
section_list = [
    "NPB450X190X77.57", "ISA100X100X10", "UC254X254X89", "NPB600X220X122.45", "UC305X305X118", 
    "ISMC100", "ISMC125", "ISMC150", "UC152X152X30", "ISMB200", "ISA90X90X10", "ISA90X90X8", 
    "ISMC200", "ISA110X110X12", "UC356X368X153", "WPB600X300X211.92", "ISA90X90X6", "ISMB300", 
    "WPB600X300X128.79", "NPB500X200X79.36", "ISMC250", "ISA50X50X6", "ISA130X130X12", 
    "ISA65X65X6", "ISMC300", "ISMC350", "UC203X203X46", "UB305X165X46", "ISMC75", "UB203X133X25", 
    "UC203X203X60", "ISA75X75X8", "UC305X305X97", "ISMB250", "ISA100X100X8", "TS32NB3.2", 
    "ISA75X75X6", "PIPE355.6*8", "D8", "UB356X171X67", "UB254X146X37", "ISA150X150X16", 
    "ISA150X150X12", "NPB500X200X90.68", "UB305X165X54", "ISMB150", "NPB400X180X66.31", 
    "UB533X210X92", "UB254X146X43", "UC152X152X23", "ISA65X65X8", "UB305X165X40", 
    "NPB450X190X67.16", "NPB400X180X57.38", "UC356X368X202", "NPB400X180X66.3", "UB203X133X30", 
    "NPB600X220X107.56", "ISMB100", "ISMB450", "UC305X305X240"
]

# E4 aur E6 ke Dropdowns 
drawing_no = st.selectbox("Select Drawing No (Excel Cell: E4):", drawing_list) 
section = st.selectbox("Select Section (Excel Cell: E6):", section_list)

# D10 aur E10 ke liye Excel jaisi Table banana
st.write("Enter Length (D) and QTY (E):")
# Starting me 5 khali rows dikhayenge (aap isko bada/ghata sakte hain)
df_empty = pd.DataFrame([{"Length": 0.0, "QTY": 0}] * 5)

# st.data_editor se aap mobile me hi directly values type kar payenge
edited_data = st.data_editor(df_empty, num_rows="dynamic", use_container_width=True)

# Submit Button
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

        # Download button taaki aap check kar sakein ki data sahi jagah gaya ya nahi
        with open(new_filename, "rb") as file:
            st.download_button(label="Download Updated Excel File", data=file, file_name=new_filename)
            
    except Exception as e:
        st.error(f"Error aayi hai: {e}. Kripya check karein ki file me 'Automation' naam ki sheet mojood hai.")
