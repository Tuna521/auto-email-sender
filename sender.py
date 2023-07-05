from appscript import app, k
from csv import reader

outlook = app('Microsoft Outlook')

template_file = open("email_template.html", "r")

gen_info_file = open("General_Info.csv")
gen_info_reader = reader(gen_info_file)
gen_info_header = []
gen_info_header = next(gen_info_reader)
print(gen_info_header)



msg = outlook.make(
    new=k.outgoing_message,
    with_properties={
        k.subject: 'Test Email',
        k.content: template_file.read()})

msg.make(
    new=k.recipient,
    with_properties={
        k.email_address: {
            k.name: 'Fake Person',
            k.address: 'fakeperson@gmail.com'}})

msg.open()
msg.activate()