from appscript import app, k
from csv import reader

# Specify which app to use
outlook = app('Microsoft Outlook')

# Get the template email from the html
template_file = open("email_template.html", "r")
general_email = template_file.read()

# Open the general info file and get the column name and the corresponding variables
gen_info_file = open("General_Info.csv")
gen_info_reader = reader(gen_info_file)
gen_info_header = next(gen_info_reader)

gen_info = next(gen_info_reader)

# Subject of the email
subject = "Test Email"

# Replace all instances of variables in email with variables from gen info file
for i in range(len(gen_info_header)):

    if (gen_info_header[i] == "Musical Name"):
        subject = gen_info[i] + " tickets!!"

    general_email = general_email.replace(
        '{' + gen_info_header[i] + '}', gen_info[i])

# Open the csv for each user and send them custom email
customer_info_file = open("Purchase_Summary_Dummy.csv")
customer_info_reader = reader(customer_info_file)
customer_info_header = next(customer_info_reader)
num_customer_header = len(customer_info_header)

for customer_info in customer_info_reader:

    custom_email = general_email
    address = 'fakeperson@gmail.com'
    name = 'Fake Person'

    for i in range(num_customer_header):

        if (customer_info_header[i] == "Email"):
            address = customer_info[i]

        if (customer_info_header[i] == "First Name"):
            name = customer_info[i]

        custom_email = custom_email.replace(
            '{' + customer_info_header[i] + '}', customer_info[i])

    msg = outlook.make(
        new=k.outgoing_message,
        with_properties={
            k.subject: subject,
            k.content: custom_email})

    msg.make(
        new=k.recipient,
        with_properties={
            k.email_address: {
                k.name: name,
                k.address: address}})

    msg.open()
    msg.activate()
