# Optimus
This utility is the first part of a 3-part management tool I called **Optimus**.

The purpose of Optimus is to streamline the logging and retrieval of submittals' and RFIs' (Request For Information) information.

Part 1: [Optimus NewForma Procore Log](https://github.com/antoine-carpentier/Optimus-NewForma-Procore-Log)  
Part 2: [AWS Optimus I](https://github.com/antoine-carpentier/AWS-Optimus-I)  
Part 3: [AWS Optimus II](https://github.com/antoine-carpentier/AWS-Optimus-II)

## Optimus NewForma Procore Log

This part of Optimus scans through the unread emails of a specified Outlook folder and looks for specific email subjects from either [Procore](https://www.procore.com) or [Newforma](https://www.newforma.com/).  
If it finds matches, it then downloads the files linked in said emails to their respective folders, logs the items into a Google Sheets spreadsheet and sends a Slack notification to the appropriate channel to infom the users of the new items.

See video demo of this part here: [Youtube Video](https://www.youtube.com/watch?v=9eONBx06qv0)
