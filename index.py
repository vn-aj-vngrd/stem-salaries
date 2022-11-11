from openpyxl import Workbook
import pandas as pd
from numpy import random


def prepareCompany(df):
    header = ("CompanyName", "Location", "Country", "City", "State")

    book = Workbook()
    sheet = book.active
    sheet.append(header)

    for index, row in df.iterrows():
        company = row[1]
        location = row[5]
        country = ""
        city = ""
        state = ""

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

        row = (
            company,
            location,
            country,
            city,
            state
        )

        print(row)
        sheet.append(row)

    book.save("data/company.xlsx")
    print("Done")


def prepareJob(df):
    header = ("JobTitle", "JobLevel", "JobTag")

    book = Workbook()
    sheet = book.active
    sheet.append(header)

    for index, row in df.iterrows():
        jobtitle = row[3]
        jobevel = row[2]
        jobTag = row[8]

        row = (
            jobtitle,
            jobevel,
            jobTag
        )

        print(row)
        sheet.append(row)

    book.save("data/job.xlsx")
    print("Done")


def prepareDemographic(df):
    header = ("DemoId", "Race", "Sex")

    book = Workbook()
    sheet = book.active
    sheet.append(header)

    for index, row in df.iterrows():
        race = row[27]
        sex = row[12]

        if (pd.isna(race)):
            race = "White"

        if (pd.isna(sex)):
            sex = random.choice(["Male", "Female"])

        demoId = race + sex

        row = (
            demoId,
            race,
            sex
        )

        print(row)
        sheet.append(row)

    book.save("data/demographic.xlsx")
    print("Done")


def prepareExperience(df):
    header = ("ExpId", "YearsAtCompany", "YearsOfEXperience",
              "Education")

    book = Workbook()
    sheet = book.active
    sheet.append(header)

    for index, row in df.iterrows():
        yearsAtCompany = round(row[7])
        yearsOfExperience = round(row[6])
        education = row[28]

        if (pd.isna(education)):
            education = random.choice(
                ["PhD", "Master's Degree", "Bachelor's Degree", "Some College", "Highschool"])

        expId = education.strip() + str(yearsAtCompany) + str(yearsOfExperience)

        row = (
            expId,
            yearsAtCompany,
            yearsOfExperience,
            education
        )

        print(row)
        sheet.append(row)

    book.save("data/experience.xlsx")
    print("Done")


def main():

    df = pd.read_csv("source_data.csv", index_col=None)

    # prepareJob(df)
    # prepareCompany(df)
    # prepareDemographic(df)
    prepareExperience(df)


main()
