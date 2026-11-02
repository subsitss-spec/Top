# NDT-B Service Portal (Maker and Supplier)

Apps Script project SetMakersheet (bound to the Contact Maker and Supplier sheet).
Pages: Home, Makers (Maker Directory), Request (Request Check Sheet).

## Workflow in Codespaces

    cd "/workspaces/Top/Maker and Supplier"
    clasp pull          # get latest code from Apps Script
    clasp push          # send local changes to Apps Script
    clasp deployments   # list deployments
    clasp deploy -i DEPLOYMENT_ID -d "V1.02"   # update the web app link

After editing, bump APP_VERSION in Code.js and the footer on every page.
Login once per Codespace: clasp login (account sabsit@n-denkei.co.th).
