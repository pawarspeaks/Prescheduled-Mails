import pandas as pd
import datetime
import smtplib

# ENTER your authentication details
GMAIL_ID = 'rajputsanket00@gmail.com'
GMAIL_PSWD = 'yourpasswordhere'
def sendEmail(to, sub, msg):
    print(f"Email to {to} sent with subject: {sub} and message {msg}")
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PSWD)

    s.sendmail(GMAIL_ID, to, f"Subject: {sub}\n\n{msg}")
    s.quit()

if __name__ == "__main__":
    df = pd.read_excel("Data.xlsx")
    # print(df)
    today = datetime.datetime.now().strftime("%d-%m")
    yearNow = datetime.datetime.now().strftime("%Y")
    # print(type(today))
    writeInd = []
    for index, item in df.iterrows():
        # print(index, item['Birthday'])
        bday = item['Birthday'].strftime("%d-%m")
        # print(bday)
        if (today == bday) and yearNow not in str(item['Year']):
            sendEmail(item['Email'], "Happy Birthday", item['Dialogue'])
            writeInd.append(index)

    print(writeInd)
    for i in writeInd:
        yr = df.loc[i, 'Year']
        df.loc[i, 'Year'] = str(yr) + ', ' +  str(yearNow)
        # print(df.loc[i, 'Year'])

    # print(df)
    df.to_excel('Data.xlsx', index=False)
