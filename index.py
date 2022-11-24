from openpyxl import Workbook
import pandas as pd
from numpy import random
from datetime import datetime
import re


def excelToText(filename):
    excelFile = pd.read_excel("./data/" + filename + ".xlsx")
    excelFile.to_csv("./text/" + filename + ".txt", index=None, header=True)


def prepareCompany(df):
    header = ("CompanyKey", "CompanyId", "CompanyName", "LocationKey")

    book = Workbook()
    sheet = book.active
    sheet.append(header)

    keys = []
    companyKey = 1
    for index, row in df.iterrows():
        companyName = str(row[1]).upper()
        country = ""
        state = ""
        city = ""
        location = row[5]
        if (location.count(",") == 1):
            country = "United States"
        elif (location.count(",") == 2):
            temp = location.split(",")
            country = temp[2].strip()
        if (location.count(",") > 0):
            temp = location.split(",")
            state = temp[1].strip()
        if (location.count(",") > 0):
            temp = location.split(",")
            city = temp[0].strip()
        locationKey = re.sub('\W', '', country.strip().upper(
        ) + state.strip().upper() + city.strip().upper())
        companyId = re.sub('\W', '', companyName + locationKey)

        data = (
            companyKey,
            companyId,
            companyName,
            locationKey,
        )

        if (companyId not in keys):
            keys.append(companyId)
            print(data)
            sheet.append(data)

            companyKey += 1

    book.save("data/company.xlsx")
    excelToText("company")
    print("Done")


def prepareLocation(df):
    header = ("LocationKey", "LocationId", "Country", "State", "CityId", "City")

    book = Workbook()
    sheet = book.active
    sheet.append(header)

    keys = []
    locationKey = 1
    for index, row in df.iterrows():
        country = ""
        state = ""
        cityId = row[14]
        city = ""

        location = row[5]

        # Country
        if (location.count(",") == 1):
            country = "United States"
        elif (location.count(",") == 2):
            temp = location.split(",")
            country = temp[2].strip()

        # State
        if (location.count(",") > 0):
            temp = location.split(",")
            state = temp[1].strip()

        # City
        if (location.count(",") > 0):
            temp = location.split(",")
            city = temp[0].strip()

        locationId = re.sub('\W', '', country.strip().upper(
        ) + state.strip().upper() + city.strip().upper())

        data = (
            locationKey,
            locationId,
            country,
            cityId,
            city,
            state,
        )

        if (locationId not in keys):
            keys.append(locationId)
            print(data)
            sheet.append(data)

            locationKey += 1

    book.save("data/location.xlsx")
    excelToText("location")
    print("Done")


def prepareJob(df):
    header = ("JobKey", "JobId", "JobTitle", "JobLevel", "Job Tag")

    book = Workbook()
    sheet = book.active
    sheet.append(header)

    keys = []
    jobKey = 1
    for index, row in df.iterrows():
        jobTitle = str(row[3])
        jobLevel = str(row[2])
        jobTag = "None" if (pd.isna(row[8])) else str(row[8])
        jobId = re.sub('\W', '', jobTitle.strip().upper(
        ) + jobLevel.strip().upper() + jobTag.strip().upper())

        data = (
            jobKey,
            jobId,
            jobTitle,
            jobLevel,
            jobTag,
        )

        if (jobId not in keys):
            print(data)
            keys.append(jobId)
            sheet.append(data)

            jobKey += 1

    book.save("data/job.xlsx")
    excelToText("job")
    print("Done")


def prepareDemographic(df):
    header = ("DemoKey", "DemoId", "Race", "Gender")

    book = Workbook()
    sheet = book.active
    sheet.append(header)

    keys = []
    demoKey = 1
    for index, row in df.iterrows():
        race = "White" if pd.isna(row[27]) else str(row[27])
        gender = random.choice(["Male", "Female"]) if pd.isna(
            row[12]) or row[12] == "Title: Senior Software Engineer" else str(row[12])
        demoId = re.sub('\W', '', race.strip().upper() +
                        gender.strip().upper())

        data = (
            demoKey,
            demoId,
            race,
            gender
        )

        if (demoId not in keys):
            keys.append(demoId)
            print(data)
            sheet.append(data)

            demoKey += 1

    book.save("data/demographic.xlsx")
    excelToText("demographic")
    print("Done")


def prepareExperience(df):
    header = ("ExpKey", "ExpId", "YearsAtCompany", "YearsOfExperience",
              "Education")

    book = Workbook()
    sheet = book.active
    sheet.append(header)

    keys = []
    expKey = 1
    for index, row in df.iterrows():
        yearsAtCompany = round(row[7])
        yearsOfExperience = round(row[6])
        education = "Highschool" if pd.isna(row[28]) else str(row[28])
        expId = re.sub('\W', '', education.strip(
        ).upper() + str(yearsAtCompany) + str(yearsOfExperience))

        data = (
            expKey,
            expId,
            yearsAtCompany,
            yearsOfExperience,
            education
        )

        if (expId not in keys):
            keys.append(expId)
            print(data)
            sheet.append(data)

            expKey += 1

    book.save("data/experience.xlsx")
    excelToText("experience")
    print("Done")


def prepareTime(df):
    header = ("TimeKey", "CalendarDateChar", "CalendarDate", "Year", "Month", "MonthName",
              "Day", "DayName", "DayOfYear")

    book = Workbook()
    sheet = book.active
    sheet.append(header)

    keys = []
    timeKey = 1
    for index, row in df.iterrows():
        calendarDate = datetime.strptime(
            row[0], '%m/%d/%Y %H:%M').date()
        year = calendarDate.year
        month = calendarDate.month
        monthName = calendarDate.strftime("%B")
        day = calendarDate.day
        dayName = calendarDate.strftime("%A")
        dayOfYear = calendarDate.strftime('%j')

        data = (
            timeKey,
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

            timeKey += 1

    book.save("data/time.xlsx")
    excelToText("time")
    print("Done")


def prepareSalary(df):
    header = ("CompanyKey", "JobKey", "DemoKey", "ExpKey", "TimeKey",
              "BaseSalary", "TotalYearlyCompensation", "Bonus")

    book = Workbook()
    sheet = book.active
    sheet.append(header)

    keys = []
    for index, row in df.iterrows():
        companyName = str(row[1]).upper()
        country = ""
        state = ""
        city = ""
        location = row[5]
        if (location.count(",") == 1):
            country = "United States"
        elif (location.count(",") == 2):
            temp = location.split(",")
            country = temp[2].strip()
        if (location.count(",") > 0):
            temp = location.split(",")
            state = temp[1].strip()
        if (location.count(",") > 0):
            temp = location.split(",")
            city = temp[0].strip()
        locationKey = re.sub('\W', '', country.strip().upper(
        ) + state.strip().upper() + city.strip().upper())
        companyKey = re.sub('\W', '', companyName + locationKey)

        jobTitle = str(row[3])
        jobLevel = str(row[2])
        jobTag = "None" if (pd.isna(row[8])) else str(row[8])
        jobKey = re.sub('\W', '', jobTitle.strip().upper(
        ) + jobLevel.strip().upper() + jobTag.strip().upper())

        race = "White" if pd.isna(row[27]) else str(row[27])
        gender = random.choice(["Male", "Female"]) if pd.isna(
            row[12]) or row[12] == "Title: Senior Software Engineer" else str(row[12])
        demoKey = re.sub('\W', '', race.strip().upper() +
                         gender.strip().upper())

        yearsAtCompany = round(row[7])
        yearsOfExperience = round(row[6])
        education = "Highschool" if pd.isna(row[28]) else str(row[28])
        expKey = re.sub('\W', '', education.strip(
        ).upper() + str(yearsAtCompany) + str(yearsOfExperience))

        timeKey = datetime.strptime(
            row[0], '%m/%d/%Y %H:%M').date()

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
    excelToText("salary")
    print("Done")


def main():
    df = pd.read_csv("source/source_data.csv", index_col=None)

    prepareCompany(df)
    prepareLocation(df)
    prepareJob(df)
    prepareDemographic(df)
    prepareExperience(df)
    prepareTime(df)
    prepareSalary(df)


main()
