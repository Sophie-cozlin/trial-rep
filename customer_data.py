import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw
excel_file_path='test_enginering.xlsx'


def main():


     df=pd.read_excel(excel_file_path ,engine='openpyxl')
     loan_data=pd.read_excel('test_enginering.xlsx',sheet_name='loan_data')
     transaction_data=pd.read_excel('test_enginering.xlsx',sheet_name='transaction_data')

     loan_data=pd.DataFrame(loan_data)
     transaction_data_df=pd.DataFrame(transaction_data)

     print(loan_data.head(10))                                                                               
     print(transaction_data.head(10)) 

     merged_df=pd.merge(loan_data,transaction_data_df,on="Customer_ID")  
     print(merged_df.head(10))
     filtered_df = merged_df[merged_df['Loan_Amount'] > 50000]
     print(filtered_df)

     # Print the filtered DataFrame
     print(filtered_df[['Loan_Type', 'Loan_Amount']])
     new_sheet_name='filteredloandata'

     with pd.ExcelWriter(excel_file_path, mode='a', engine='openpyxl') as writer:
          filtered_df[['Loan_Type', 'Loan_Amount']].to_excel(writer,sheet_name=new_sheet_name, index=False)
          
     print("data has been successfully printed")
     
     new_sheet_name='FilteredData'
     with pd.ExcelWriter(excel_file_path, mode='a', engine='openpyxl') as writer:    
     
          filtered_df.to_excel(writer,sheet_name=new_sheet_name, index=False)
     print("data has been successfully printed")

     # new_sheet_data =  filtered_df[['Loan_Type', 'Loan_Amount']][ filtered_df[['Loan_Type', 'Loan_Amount']]['Sheet Name'] == 'loans']
     filtered_df.sort_values(by='Loan_Type', inplace=True)
     # Step 3: Plot the graph
     plt.figure(figsize=(10, 6))  # Optional: Adjust the figure size
     plt.bar(filtered_df['Loan_Type'], filtered_df['Loan_Amount'])
     plt.xlabel('Loan Type')
     plt.ylabel('Loan Amount')
     plt.title('Loan Amount vs Loan Type for High Value Loans')
     plt.xticks(rotation=45)  # Rotate x-axis labels for better readability
     plt.tight_layout()  # Adjust layout to prevent overlapping labels
     #plt.show()
     new_sheet_data =  filtered_df[['Loan_Type', 'Loan_Amount']][ filtered_df[['Loan_Type', 'Loan_Amount']]['Sheet Name'] == 'loans']
     print("data plotted successfully")
   
   
# '__name__' == '__main__':
main()