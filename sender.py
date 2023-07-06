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
gen_info_header = []
gen_info_header = next(gen_info_reader)

gen_info = next(gen_info_reader)

# Replace all instances of variables in email with variables from gen info file
for i in range(len(gen_info_header)):
    general_email = general_email.replace(
        '{' + gen_info_header[i] + '}', gen_info[i])

# Open the csv for each user and send them custom email
custom_email = general_email

customer_info_file = open("Purchase_Summary_Dummy.csv")

# for i in range(len())

msg = outlook.make(
    new=k.outgoing_message,
    with_properties={
        k.subject: 'Test Email',
        k.content: general_email})

msg.make(
    new=k.recipient,
    with_properties={
        k.email_address: {
            k.name: 'Fake Person',
            k.address: 'fakeperson@gmail.com'}})

msg.open()
msg.activate()
