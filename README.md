# dash_app_proposal_info_Elsevier_Expert_Lookup
Funding organizations who use Elsevier Expert Lookup for searching reviewers for subsidy applications manually place proposals information per subsidy round in the Expert Lookup. That takes a lot of time and energy. By automatically placing the proposals information in the Expert Lookup, funding organizations can save a lot of time.

Elsevier Expert Lookup provides format in which it can take the lumsum proposals information. Here under written code provides the missing gap, where the lumsum proposals information from funding organization is converted into the required format to upload in Expert Lookup.

In preparation, the number of files that you must download and edit in Excel in advance are:

Example list proposals information from expert Lookup ‘Proposal Upload Example.xlsx’ :

a. Put worksheets ‘Suggested reviewers’ and ‘country sheet’ from ‘Proposal Upload Example.xlsx’ as a separate Excel files

b. Manually fill the ‘Suggested reviewers’ Excel with information, if available for the entire set of proposals

The code creates web user interface via Dash Application. If you run this code in IDE, you can copy past ‘http://127.0.0.1:8050/' in your browser to start uploading the file with proposal information. An example of such a file columns is given. The output will be a excel file named:’Proposal info to feed in EL.xlsx’
# Check proposals information Excel
This app produces two tabs (proposal details and applicant information) from the example ‘Proposal Upload Example.xlsx’. Add in the excel two more Put worksheets ‘Suggested reviewers’ and ‘country sheet’. Check if the excel file columns and format is same as the example.

# Upload proposals information in Expert Lookup
In the environment of workspace of Export Lookup, use ‘import proposals’ to upload proposals information Excel ‘Proposal info to feed in EL’. If you get an error message, check the form of the Excel file again. If the problem still persists, copy the information from ‘Proposal info to feed in EL.xlsx’ and paste in ‘Proposal Upload Example.xlsx’. Then upload ‘Proposal Upload Example.xlsx’ in expert Lookup. It takes some time to upload all the information. The Expert lookup shows a note when it’s done.
## Code
This is a code for a Dash app that allows a user to upload a CSV or Excel file, parse the contents, and transform the data into a specific format. The app uses various libraries, including pandas, re, and dash, to manipulate and display the data. The user can drag and drop or select a file to upload, and the app will then display the transformed data in the output section. The specific format for the transformed data is described in the function parse_contents, which includes the creation of two dataframes, df_PD and df_PA, and the manipulation of various columns within those dataframes.
