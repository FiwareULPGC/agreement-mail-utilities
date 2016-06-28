# Agreement mail utilities

This repository contains scripts to manage the agreement emails processing.

## Python dependencies

The script depends on the following Python packages

* [PyYAML](https://pypi.python.org/pypi/PyYAML)
* [openpyxl](https://pypi.python.org/pypi/openpyxl/2.3.3)

They can be easily installed via pip:

``` sudo pip install pyyaml openpyxl```

## Configure Gmail

These scripts have been tested using Gmail as the SMTP server. In order to 
enable your chosen Gmail account to allow third party apps to log in and send 
emails remotely, you should:

* Log in to your Gmail account.
* Access to https://www.google.com/settings/security/lesssecureapps.
* Choose **Turn on**.

From that point on, you should be able to send emails specifying the account
address and password in the configuration file.

## generate_confirmation_tokens.py
 
The script takes a list of names and emails from an xlsx document and 
generates a uniquely identified token per user. Each token will be saved back 
to the input xlsx.

The format expected for the input file is:
```
Name | Surname | email
```

The file will then be extended with two extra columns:
```
Name | Surname | email | Token | Sent
```

Every time the script is executed, a new token is generated for each row and 
the 'Sent' column is set to 'No'.

**Example of use**:
```
python ./generate_confirmation_tokens.py -f ./<filename>.xlsx
```


## send_confirmation_email.py
The script takes the previously generated xlsx file which containts a list of
names, emails and tokens, and a configuration file to send a personalized email 
for each entry.

The configuration file has the following fields:

* **smtp_user** and **smtp_pass** are the credentials for the Gmail account
  from where the agreement mails will be sent.
* **smtp_host** is the URL for the SMTP server from where the mails will sent.
* **subject** is the value that will appear in the field SUBJECT.
* **template** is the body of the email to be sent. You can add the special
  strings <name> <surname> and <token>, which will be sustituted with their 
  respective cell values from the input xlsx file.

If the email generation and and request are succesful, the row will be marked
as sent by writing 'Yes' in the 'Sent' column. Otherwise, 'No' will be written.
You can also force to send mails already marked as sent by indicating the resend
value to ```True```.


**Example of default use**:
```
python ./send_confirmation_email.py -c <configfilename>.yml -f ./<filename>.xlsx
```

**If you want to force a resend for every row then execute as**
```
python ./send_confirmation_email.py -c <configfilename>.yml -f ./<filename>.xlsx --resend
```
