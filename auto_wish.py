import pandas as pd
import datetime
import smtplib

# Your Detail

GMAIL_ID = "yourgmail@gmail.com"
GMAIL_PSWD = "your password"  #ENTER YOUR GMAIL PASSWORD


def sendEmail(to, sub, msg):
    print(f"A {sub} Email sent to {to} and a message : {msg}")
    s = smtplib.SMTP("smtp.gmail.com", 587)
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PSWD)
    s.sendmail(GMAIL_ID, to, f"Subject: {sub}\n\n{msg}")
    s.quit()


if __name__ == '__main__':

    df = pd.read_excel(f"data.xlsx")
    df = df.astype(str)
    # print(df)
    today = datetime.datetime.now().strftime("%d-%m")
    this_year = datetime.datetime.now().strftime("%Y")
    # print(this_year)
    # print(today)

    writeInd = []
    for index, item in df.iterrows():
        # print(index, item)
        bday = item["Birthday"]
        # print(bday)

        if today == bday and this_year not in item["Year"]:
            sendEmail(item['Email'], "Happy BirthDay", item["Dialogue"])
            writeInd.append(index)

    print(writeInd)
    for i in writeInd:
        yr = df.loc[i, "Year"]
        # print(yr)
        df.loc[i, "Year"] = yr + "," + this_year
        # print(df.loc[i, "Year"])

    # print(df)
    df.to_excel("data.xlsx", index=False)
