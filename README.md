# GSheets

This demonstrates how to use GRAX Data APIs to quickly reuse data from Salesforce in Google Sheets for business intelligence.

To use:

- [Get GRAX](https://grax.com), [connect to Salesforce](https://documentation.grax.com/docs/connecting-salesforce), and [start the data collector](https://documentation.grax.com/docs/auto-backup)
- Log into the web app and generate a [public API key](https://documentation.grax.com/docs/public-api)
- Clone [Sample Google Sheet](https://docs.google.com/spreadsheets/d/1MbdocT6b1sB65HhmepzTgpNxCcg0KcyICZ5U6sPz_t4/edit#gid=514958285)
- Edit Cloned Google Document App Scripts
    - Click ```Menu Extensions```
    - Select ```Apps Scripts```
        - Update APIName example ```APIName='https://<YOUR-APP-NAME>.secure.grax.io/';```
        - Update APIToken example ```APIToken=grax_token_Rne96drL5FEW1SAMPLesampleTokenSaMPleEdit9KQF;```
        - Click Save Icon in menu bar (Disc Image)
- Click ```GRAX Data``` in Menu Bar
    - Click ```Run Demo```
    - Allow Scripts to run
- Click ```GRAX Data``` in Menu Bar (Again)
    - Click ```Run Demo```
- Click Run and see your Google Sheet populated with the Historical Opportunity data

