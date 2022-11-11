from openpyxl import Workbook
import pandas as pd
from numpy import random
from datetime import datetime


def prepareLocation(df):
    header = ("CompanyKey", "Location", "Country",
              "CityId", "City", "State", "DMAId")

    book = Workbook()
    sheet = book.active
    sheet.append(header)

    keys = []
    for index, row in df.iterrows():
        if (row[5] not in keys):
            companyKey = str(row[1]).upper()
            location = row[5]
            country = ""
            cityId = row[14]
            city = ""
            state = ""
            dmaId = row[15]

            # Country
            if (location.count(",") == 1):
                country = "United States"
            elif (location.count(",") == 2):
                temp = location.split(",")
                country = temp[2].strip()

            # City
            if (location.count(",") > 0):
                temp = location.split(",")
                city = temp[0].strip()

            # State
            if (location.count(",") > 0):
                temp = location.split(",")
                state = temp[1].strip()

            data = (
                companyKey,
                location,
                country,
                cityId,
                city,
                state,
                dmaId,
            )

            keys.append(row[5])
            print(data)
            sheet.append(data)

    book.save("data/location.xlsx")
    print("Done")


def prepareCompany(df):
    header = ("CompanyName", )

    book = Workbook()
    sheet = book.active
    sheet.append(header)

    keys = []
    for index, row in df.iterrows():
        if (str(row[1]).upper() not in keys):
            companyName = str(row[1]).upper()

            data = (
                companyName,
            )

            keys.append(companyName)
            print(data)
            sheet.append(data)

    book.save("data/company.xlsx")
    print("Done")


def prepareJob(df):
    header = ("JobId", "JobTitle", "JobLevel", "JobTag")

    book = Workbook()
    sheet = book.active
    sheet.append(header)

    keys = []
    for index, row in df.iterrows():
        jobTitle = str(row[3])
        jobLevel = str(row[2])
        jobTag = "None" if pd.isna(row[8]) else str(row[8])
        jobId = jobTitle.replace(
            " ", "").strip() + jobLevel.replace(" ", "").strip()

        data = (
            jobId,
            jobTitle,
            jobLevel,
            jobTag
        )

        if (jobId not in keys):
            print(data)
            keys.append(jobId)
            sheet.append(data)

    book.save("data/job.xlsx")
    print("Done")


def prepareDemographic(df):
    header = ("DemoId", "Race", "Gender")

    book = Workbook()
    sheet = book.active
    sheet.append(header)

    keys = []
    for index, row in df.iterrows():
        race = "White" if pd.isna(row[27]) else str(row[27])
        gender = random.choice(["Male", "Female"]) if pd.isna(
            row[12]) or row[12] == "Title: Senior Software Engineer" else str(row[12])

        demoId = race.replace(" ", "").strip() + gender

        data = (
            demoId,
            race,
            gender
        )

        if (demoId not in keys):
            keys.append(demoId)
            print(data)
            sheet.append(data)

    book.save("data/demographic.xlsx")
    print("Done")


def prepareExperience(df):
    header = ("ExpId", "YearsAtCompany", "YearsOfExperience",
              "Education")

    book = Workbook()
    sheet = book.active
    sheet.append(header)

    keys = []
    for index, row in df.iterrows():
        yearsAtCompany = round(row[7])
        yearsOfExperience = round(row[6])
        education = random.choice(
            ["PhD", "Master's Degree", "Bachelor's Degree", "Some College", "Highschool"]) if pd.isna(row[28]) else str(row[28])

        expId = education.replace(" ", "").strip(
        ) + str(yearsAtCompany) + str(yearsOfExperience)

        data = (
            expId,
            yearsAtCompany,
            yearsOfExperience,
            education
        )

        if (expId not in keys):
            keys.append(expId)
            print(data)
            sheet.append(data)

    book.save("data/experience.xlsx")
    print("Done")


def prepareTime(df):
    header = ("CalendarDateChar", "CalendarDate", "Year", "Month", "MonthName",
              "Day", "DayName", "DayOfYear")

    book = Workbook()
    sheet = book.active
    sheet.append(header)

    keys = []
    for index, row in df.iterrows():
        calendarDate = datetime.strptime(
            row[0], '%m/%d/%Y %H:%M:%S').date()
        year = calendarDate.year
        month = calendarDate.month
        monthName = calendarDate.strftime("%B")
        day = calendarDate.day
        dayName = calendarDate.strftime("%A")
        dayOfYear = calendarDate.strftime('%j')

        data = (
            str(calendarDate),
            calendarDate,
            year,
            month,
            monthName,
            day,
            dayName,
            dayOfYear,
        )

        if (str(calendarDate) not in keys):
            keys.append(str(calendarDate))
            print(data)
            sheet.append(data)

    book.save("data/time.xlsx")
    print("Done")


def prepareSalary(df):
    header = ("CompanyKey", "JobKey", "DemoKey", "ExpKey", "TimeKey",
              "BaseSalary", "TotalYearlyCompensation", "Bonus")

    book = Workbook()
    sheet = book.active
    sheet.append(header)

    keys = []
    for index, row in df.iterrows():
        companyKey = str(row[1]).upper()
        jobKey = str(row[3]).replace(
            " ", "").strip() + str(row[2]).replace(" ", "").strip()

        race = "White" if pd.isna(row[27]) else str(row[27])
        gender = random.choice(["Male", "Female"]) if pd.isna(
            row[12]) or row[12] == "Title: Senior Software Engineer" else str(row[12])
        demoKey = race.replace(" ", "").strip() + gender

        yearsAtCompany = round(row[7])
        yearsOfExperience = round(row[6])
        education = random.choice(
            ["PhD", "Master's Degree", "Bachelor's Degree", "Some College", "Highschool"]) if pd.isna(row[28]) else str(row[28])
        expKey = education.replace(" ", "").strip(
        ) + str(yearsAtCompany) + str(yearsOfExperience)

        timeKey = datetime.strptime(
            row[0], '%m/%d/%Y %H:%M:%S').date()

        baseSalary = 1000 if pd.isna(row[9]) or row[9] == 0 else row[9]
        totalYearlyCompensation = row[4]
        bonus = row[11]

        data = (
            companyKey,
            jobKey,
            demoKey,
            expKey,
            timeKey,
            baseSalary,
            totalYearlyCompensation,
            bonus,
        )

        print(data)
        sheet.append(data)

    book.save("data/salary.xlsx")
    print("Done")


def main():
    df = pd.read_csv("source/source_data.csv", index_col=None)

    prepareLocation(df)
    prepareCompany(df)
    prepareJob(df)
    prepareDemographic(df)
    prepareExperience(df)
    prepareTime(df)
    prepareSalary(df)


main()
