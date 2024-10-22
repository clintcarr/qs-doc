# About
QS:Doc is a standalone application that ingests information from the Qlik Sense Repository Service and creates a both a formatted Microsoft word document and a Microsoft Office Excel spreadsheet containing the same information but not formatted (for easy ingestion into an application or database)

The output can then be used as a build document, disaster recovery document or purely as information on configuration.

# What it collects

| Site Details | Assets | Hardware | Properties |
|--------------|--------|----------|------------|
| Service Cluster | Applications | Windows Management information | Custom Properties
| Engine information | Extensions | | Tags
| Proxy information | Data Connections | |User Directories
| Virtual Proxy information | System Rules || Site License
| Scheduler information | License Rules
| Repository information




# Versions

Ensure the correct version is used for the version of Qlik Sense Enterprise that is being documented.

| Qlik Sense Version | QS:Doc Version |
|--------------------|----------------|
| 3.2.x | QS:Doc 3.2|
| 3.1.6 | QS:Doc 3.2|
| 3.1.0 | QS:Doc 3.0|
| 3.0.x | QS:Doc 3.0|

## Installation
1. Ensure the c runtime is installed (this is usually installed via Windows update - so there should be no need to install this separately for QS:Doc)
https://support.microsoft.com/en-au/help/2999226/update-for-universal-c-runtime-in-windows
2. Copy the folder (qliksense3.0 or qliksense3.2) to a local location.  Ensure the word_template folder contains the word document.
3. Ensure Qlik Sense Port 4242 is open if using certificate authentication, else 443/4244 (80/4248) if using Windows Authentication

## Usage
Launch .\Create_QSDoc.exe --help to view the usage.

![alt text](https://github.com/clintcarr/qs-doc/blob/master/help.png)

1. To connect using Windows Authentication use the user and password switches: .\Create_QSDoc.exe --server {servername} --user {domain\username} --password {password}
2. To connect using certificate authentication ensure the certificates are exported as PEM format: .\Create_QSDoc.exe --server {servername} --certs {path to PEM files}
3. To capture WMI information ensure you use Windows Authentication (captures RAM/CPU information)

### Windows Authentication
![alt text](https://github.com/clintcarr/qs-doc/blob/master/capture.gif)
