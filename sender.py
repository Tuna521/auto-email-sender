from appscript import app, k

outlook = app('Microsoft Outlook')

f = open("email_template.html", "r")

msg = outlook.make(
    new=k.outgoing_message,
    with_properties={
        k.subject: 'Test Email',
        k.content: f.read()})

msg.make(
    new=k.recipient,
    with_properties={
        k.email_address: {
            k.name: 'Fake Person',
            k.address: 'fakeperson@gmail.com'}})

msg.open()
msg.activate()