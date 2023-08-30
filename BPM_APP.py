import pandas as pd
import numpy as np
import os
import warnings
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from datetime import datetime
import tempfile
import win32com.client as win32
import pythoncom
st.set_page_config(initial_sidebar_state="expanded")

# Here provide the excel data as null
excel_data = None

# Function to download Excel data
def download_excel(dataframes):
    op = BytesIO()
    with pd.ExcelWriter(op, engine='xlsxwriter') as wr:
        for sheet_name, df in dataframes.items():
            df.to_excel(wr, index=False, sheet_name=sheet_name)
            workbook = wr.book
            worksheet = wr.sheets[sheet_name]
            # 1 
            if sheet_name == 'Applicant_wise_count_FY':
                fm = workbook.add_format({'num_format': '0.00'})
                worksheet.set_column("A:A", None, fm)
            else:
                fm = workbook.add_format({'bold': True})
                worksheet.set_column("A:A", None, fm)
            # 2    
            if sheet_name == 'Total_Amount_applicant_wise_FY':
                fm = workbook.add_format({'num_format': '0.00'})
                worksheet.set_column("A:A", None, fm)
            else:
                fm = workbook.add_format({'bold': True})
                worksheet.set_column("A:A", None, fm)  
                
            #3
            if sheet_name == 'department_wise_count_Days':
                fm = workbook.add_format({'num_format': '0.00'})
                worksheet.set_column("A:A", None, fm)
            else:
                fm = workbook.add_format({'bold': True})
                worksheet.set_column("A:A", None, fm) 
                
            #4
            
            if sheet_name == 'Applicant_Wise_Count_Days':
                fm = workbook.add_format({'num_format': '0.00'})
                worksheet.set_column("A:A", None, fm)
            else:
                fm = workbook.add_format({'bold': True})
                worksheet.set_column("A:A", None, fm)   
                
            #5
            
            if sheet_name == 'Dump':
                fm = workbook.add_format({'num_format': '0.00'})
                worksheet.set_column("A:A", None, fm)
            else:
                fm = workbook.add_format({'bold': True})
                worksheet.set_column("A:A", None, fm)          
    op.seek(0)
    return op.getvalue()

def send_email(email_addresses , pivot_table_1 , pivot_table_2, pivot_table_3, pivot_table_4, data):
    # Initialize the COM library
    pythoncom.CoInitialize()

    try:
        excel_io = BytesIO()
        with pd.ExcelWriter(excel_io, engine='xlsxwriter') as wr:
            pivot_table_1.to_excel(wr, sheet_name='Applicant_wise_count_FY', index=False)
            pivot_table_2.to_excel(wr, sheet_name='Total_Amount_applicant_wise_FY', index=False)
            pivot_table_3.to_excel(wr, sheet_name='department_wise_count_Days', index=False)
            pivot_table_4.to_excel(wr, sheet_name='Applicant_Wise_Count_Days', index=False)
            data.to_excel(wr, sheet_name='Dump', index=False)

        # Reset the buffer position for reading
        excel_io.seek(0)

        # Outlook constants
        olMailItem = 0
        olFormatHTML = 2

        # Create an Outlook instance
        outlook = win32.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace("MAPI")

        # Compose and send emails
        mail = outlook.CreateItem(olMailItem)
        current_date = datetime.now().strftime("%d%m%Y")

        mail.Subject = f"BPM Approval Pending cases {current_date}"
        mail.BodyFormat = olFormatHTML
        
        pivot_table_4 = pivot_table_4
        table_style = '''
        <style>
            table {

                border-collapse: collapse;
                width: 75%;
            }

            th, td {
            border: 1.5px solid black;
            padding: 8px;
            text-align: left;
            font-weight: bold;
            }
            th {
                background-color: #f2f2f2;
            }
        </style>

        '''        
        
        df_html = table_style + pivot_table_4.to_html(index=False)

        mail.HTMLBody = f'''
        <html>
        <body>
        <p>Hi Team,</p>

        <p>Please find the below BPM Approval pending cases which have crossed more than 07 days. Please check your respective department BPM and clear the approvals by today end of day (EOD) and confirm.</p>
        <p>Currently from the system data which is at applicant level in that we don’t have a clarity for BPM Approval matrix. That’s the reason we are sending an regular E-mail to Applicant so that they can follow up with respective approver’s as it’s causing delay in vendor payment. </p>
        <ul>
            <li>Approved from your end but final approval pending – Please check in your BPM for the details of whom it is pending with. Kindly follow up with the concerned person and clear the approvals by today EOD.</li>

            <li>If any applicant has left the organization, please update for re-submission.</li>

            <li>If you want to reject/cancel, please do reject and invalidate the request. Kindly confirm the same for cancelling the invoice.</li>
        </ul>

        <h3>Pending approval Summary :-:</h3>
        {pivot_table_4.to_html(index=False)}<br><br>

        <b><i>Thanks &amp; Regards<br>
                        Sakshi Garg <br>
                        <span style="color: #D2691E;">Xiaomi Technology India Private Limited</span><br>
                        <span style="color: #1E90FF;">Building Orchid. Block E, Embassy Tech Village<br>
                        Marathahalli Outer Ring Road, Deverabisanahalli, Bangaluru 560103</span></i></b><br><br>
        </body>
        </html>
        '''
        mail.CC = 'nilatpal@xiaomi.com;rvenu@xiaomi.com;v-prathmeshm1@xiaomi.com; sakshi3@xiaomi.com'
        bcc = ''
        for email in email_addresses:
            bcc = bcc + email + ';'
        bcc = bcc + 'rohit.kaushik@quation.in;tanmay@xiaomi.com;jvikas@xiaomi.com'
        mail.BCC = bcc
        
        # Attach the Excel data
        today_date = datetime.now().strftime("%Y-%m-%d")
        filename = f"BPM_{today_date}.xlsx"
        tmpfile_path = os.path.join(tempfile.gettempdir(), filename)
        with open(tmpfile_path, 'wb') as tmpfile:
            tmpfile.write(excel_io.read())
        mail.Attachments.Add(tmpfile_path)

        # Display the email for manual sending
        mail.Display()

    finally:
        # Uninitialize the COM library
        pythoncom.CoUninitialize()
def main():
    global pivot_table_4, pivot_table_1, pivot_table_2, pivot_table_3, Data, excel_data
    st.title("BPM AGEING REPORT")
    email_addresses=['r14101996@gmail.com','r14101996@gmail.com']

    uploaded_data = st.file_uploader("Upload Data File", type=["xlsx"])
    uploaded_master = st.file_uploader("Upload Master File", type=["xlsx"])

    if uploaded_data and uploaded_master:
        Master = pd.read_excel(uploaded_master, sheet_name='Sheet1')
        Data = pd.read_excel(uploaded_data, sheet_name='Sheet1')
        
        Data.drop(['Invoice No', 'Batch','Unnamed: 10','Account Due Date','Mi Xin Billing Date','Emergency Payment Reason','Comments','Batch Reason', 'Payment Account','Payment Method.1', 'Payment channel', 'Payment Status','Bank Reference No', 'Payment Date','Payment Curr.',
        'Agency','Vendor Address','Account Name(EN)','Account Address','Cashier', 'Accountant','Factoring Status', 'Assignment', 'Document No'],axis=1,inplace=True)


        #Insert the Important columns
        Data.insert(5,'Year',"FY'23-24")
        Data.insert(6,'Today',pd.Timestamp.today())
        Data['Today'] = Data['Today'].astype('datetime64[ns]')
        Data.insert(7,'Days',(pd.Timestamp.today() - Data['Appl. Date']).dt.days)
    

        #Create the Bins
        bins = [0, 7, 30, 60, 90, 180, Data['Days'].max()]
        labels = ['0-7', '8-30', '31-60', '61-90', '91-180','180+']
        Data.insert(8, 'Ageing', pd.cut(Data['Days'], bins=bins, labels=labels))

        #Create the conditions
        condition_1 = (Data['Appl. Date'] >= '2021-03-01') & (Data['Appl. Date'] <= '2022-03-31')
        condition_2 = (Data['Appl. Date'] >= '2022-03-01') & (Data['Appl. Date'] <= '2023-03-31')
        condition_3 = (Data['Appl. Date'] >= '2023-03-01') & (Data['Appl. Date'] <= '2024-03-31')
        
        # Apply fiscal year conditions
        Data.loc[condition_1, 'Year'] = "FY'21-22"
        Data.loc[condition_2, 'Year'] = "FY'22-23"
        Data.loc[condition_3, 'Year'] = "FY'23-24"
        
        #Filter the AP-Pending Approval Cases
        Data=Data[Data['Status']=='AP-Pending Approval']
        Data

        #Create the first Pivot
        pivot_table_data1 = pd.pivot_table(Data, index=['Appl. Dept.', 'Applicant'], columns='Year', values='Appl. Date', aggfunc='count')
        pivot_table_data2 = pd.pivot_table(Data, index=['Appl. Dept.'], columns='Year', values='Appl. Date', aggfunc='count', fill_value=0)
        pivot_table_data_reset_1 = pivot_table_data1.reset_index()
        pivot_table_data_reset_2 = pivot_table_data2.reset_index()
        pivot_table1 = pd.concat([pivot_table_data_reset_1,pivot_table_data_reset_2])
        new1 = pivot_table1.reset_index(drop = True)
        new1['Applicant'] = new1['Applicant'].fillna('')

        new1['''FY'23-24'''] = new1['''FY'23-24'''].fillna(0)
        new2 = new1.sort_values(by = ['Appl. Dept.', 'Applicant'],ascending =[True, True] )
        new2.loc[new1['Applicant']!='','Appl. Dept.'] = ''
        new2['Grand Total'] = new1['''FY'23-24''']
        pivot_table_1=new2
        pivot_table_1
        
        
        #Create the second Pivot
        pivot_table_data3 = pd.pivot_table(Data, index=['Appl. Dept.', 'Applicant'], columns='Year', values='Batch Amount', aggfunc='sum')
        pivot_table_data4 = pd.pivot_table(Data, index=['Appl. Dept.'], columns='Year', values='Batch Amount', aggfunc='sum', fill_value=0)
        pivot_table_data_reset_3 = pivot_table_data3.reset_index()
        pivot_table_data_reset_4 = pivot_table_data4.reset_index()

        pivot_table2 = pd.concat([pivot_table_data_reset_3,pivot_table_data_reset_4])
        new3 = pivot_table2.reset_index(drop = True)
        pd.set_option('display.float_format', '{:.2f}'.format)
        new3['Applicant'] = new3['Applicant'].fillna('')
        # new3['''FY'22-23'''] = new3['''FY'22-23'''].fillna(0)
        new3['''FY'23-24'''] = new3['''FY'23-24'''].fillna(0)
        new4 = new3.sort_values(by = ['Appl. Dept.', 'Applicant'],ascending =[True, True] )
        new4.loc[new3['Applicant']!='','Appl. Dept.'] = ''
        new4['Total'] = new4['''FY'23-24''']
        pivot_table_2=new4
        pivot_table_2



        #Create the third Pivot
        pivot_table_data_3 = pd.pivot_table(Data, index=['Appl. Dept.'], columns='Ageing', values='Application No', aggfunc='count', fill_value=0, margins=True, margins_name='Grand Total')
        pivot_table_data_3.columns = ['0-7', '8-30', '31-60', '61-90', '91-180','180+', 'Grand Total']
        pivot_table_data_3 = pivot_table_data_3.reset_index()
        pivot_table_3=pivot_table_data_3
        pivot_table_data_3
        
        
        #Create the forth Pivot
        pivot_table_data5 = pd.pivot_table(Data, index=['Appl. Dept.', 'Applicant'], columns='Ageing', values='Application No', aggfunc='count')
        pivot_table_data6 = pd.pivot_table(Data, index=['Appl. Dept.'], columns='Ageing', values='Application No', aggfunc='count', fill_value=0)
        pivot_table_data_reset_5 = pivot_table_data5.reset_index()
        pivot_table_data_reset_6 = pivot_table_data6.reset_index()
        pivot_table4 = pd.concat([pivot_table_data_reset_5,pivot_table_data_reset_6])
        new5 = pivot_table4.reset_index(drop = True)
        new5['Applicant'] = new5['Applicant'].fillna('')
        new6 = new5.sort_values(by = ['Appl. Dept.', 'Applicant'],ascending =[True, True] )
        new6.loc[new6['Applicant'] != '', 'Appl. Dept.'] = ''
        new6 = new6.loc[(new6[['0-7', '8-30', '31-60', '61-90', '91-180', '180+']] != 0).any(axis=1)]
        new6['Grand Total']= new6['0-7']+new6['8-30']+new6['31-60']+new6['61-90']+ new6['91-180']+ new6['180+']
        pivot_table_4=new6
        pivot_table_4 = pivot_table_4[['Appl. Dept.', 'Applicant','0-7', '8-30', '31-60', '61-90', '91-180', '180+', 'Grand Total']]
        pivot_table_4.columns.name = None
        pivot_table_4
        
        #Filter Records of E-mail
        
        Data.rename(columns={'Applicant':'Name'},inplace=True)
        Master=Master[['Name','Email id']]
        final_df = pd.merge(Data, Master, on='Name', how='inner')
        final_df_1=final_df[['Email id']]
        df_mail=final_df_1[['Email id']].drop_duplicates()
        df = df_mail
        email_addresses = df['Email id'].tolist()
        email_addresses        
        

        if st.button("Download Excel"):
            dataframes = {
                'Applicant_wise_count_FY': pivot_table_1, 'Total_Amount_applicant_wise_FY': pivot_table_2,
                'department_wise_count_Days': pivot_table_3, 'Applicant_Wise_Count_Days': pivot_table_4, 'Dump': Data
            }
            excel_data = download_excel(dataframes)
            st.session_state.excel_data = excel_data

            if excel_data is not None:
                st.download_button("Download Result", data=excel_data, file_name="result.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.warning("Please generate Excel data before downloading.")

        if st.button("Send Email"):
            if hasattr(st.session_state, 'excel_data') and st.session_state.excel_data is not None:
                send_email(email_addresses, pivot_table_4, pivot_table_1, pivot_table_2, pivot_table_3, Data)
            else:
                st.warning("Please generate Excel data before sending emails.")

if __name__ == "__main__":
    main()
