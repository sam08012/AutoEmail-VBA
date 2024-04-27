# AutoEmail-VBA
### Automate Sending Emails using Excel VBA

A macros based excel solution to send automatic mails to companies and HRs. 

This is a VBA/macro based solution for sending multiple emails with resume attachments, to comapanies and HRs from open source databases. Using till tool you can simply download or make a excel shett containing company's represntative name with their email address, a subject and a body to sesnf them. After pressinf the "Hire me" button in excel file you can autosend all those emails the respective subject and body. for specilized subject and body you can add it in that specific row for these columns.

Below are the steps to follow to perform to make this solution work for you.
This guide will help you set up a VBA/macro-based solution in Excel to send multiple emails to companies and HRs using open source databases.

## Prerequisites

- Microsoft Excel installed on your computer.
- Developers option is on in your excel. Just a google search away. 
- Outlook email client set up with an active email account.

## Steps

1. **Prepare Your Excel Sheet**
   - Create or download an Excel sheet containing the following columns:
     - First Name
     - Email Address
     - Subject (Note: Add multiple if you want specialized subjects for certain emails else one will do fine)
     - Body (Note: Add multiple if you want specialized bodies for certain emails)
     - Attachment (Note: Add if you want to attach files(CVs and resumes) to the emails)
   - Fill in the data for each company's representative, including their first name, email address, and any additional information like subject, body, or attachment paths.

2. **Download the Repository**
   - Clone or download this repository containing the VBA code and Excel sheet.

3. **Open Excel and the VBA Editor**
   - Open the Excel sheet provided in the repository.
   - Press `Alt + F11` to open the Visual Basic for Applications (VBA) editor.

4. **Insert VBA Code**
   - In the VBA editor, open the module provided in the repository.
   - Copy the existing VBA code from the repository and paste it into the module.

5. **Update the Macro**
   - Update the VBA code with your personalized salutations, signatures, and any other necessary modifications. All marked in code comments.
   - Ensure that the code is fetching data from the correct columns in your Excel sheet.

6. **Assign Macro to Button**
   - In Excel, we can insert a button using the provided Excel sheet. Use the already provided one or create a new one if you don't like mine.
   - To Assign the macro you created to the button. Right-click on it, select `Assign Macro`, and choosing the macro from the list.

    ![Assigned macro to "Hire me" button ](https://github.com/sam08012/AutoEmail-VBA/blob/main/Screenshot%202024-04-27%20162135.png)

7. **Test the Solution**
   - Save your Excel sheet with the macro-enabled format (`.xlsm`).
   - Fill in the necessary data in your Excel sheet.
   - Click on the "Hire Me" button to trigger the macro.
   - Verify that the emails are sent correctly from your Outlook client.
   ![Testing mails](https://github.com/sam08012/AutoEmail-VBA/blob/main/Screenshot%202024-04-27%20162228.png)

8. **Customize Further (Optional)**
   - If needed, customize the VBA code to add more functionalities, error handling, or enhancements to suit your requirements.

## Additional Notes

- Ensure that Outlook is set up properly with an active email account.
- Test the solution with a small batch of emails initially to ensure everything works as expected.
- Double-check the email addresses and content before sending emails in bulk to avoid any mistakes.
