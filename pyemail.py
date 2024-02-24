import yagmail

user = 'atharvachoudhari.truefunds@gmail.com'
app_password = 'urnp mvgv rnei vgrr'
to = 'shantanudarak.truefunds@gmail.com'
subject = 'Qualitative data analysis'
contents = ['Qualitative data analysis', 'near_52week.xlsx']

with yagmail.SMTP(user, app_password) as yag:
    yag.send(to, subject, contents)
    print('Sent email successfully')
