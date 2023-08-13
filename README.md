# Auto_Email_Sender

This repository creates emails to send to the people on the spreadsheet, based on given email template, additionally filling in the information from the spreadsheet when needed.

### Built With

- Python
- appscript library
- csv library

## Getting Started

To get a local copy up and running follow these simple example steps.

### Prerequisites

- appscript

```sh
pip install appscript
```

### Installation

1. Clone the repo

```ssh
git clone https://github.com/Tuna521/auto-email-sender.git
```

2. In the folder write template for your email content in html file `email_template.html`

3. Add the csv file with the csv information of the customers

Important: it need to contain column with "First Name" and "Quantity"
