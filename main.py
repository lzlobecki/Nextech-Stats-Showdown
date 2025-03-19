import pandas as pd
import numpy as np
from sklearn.ensemble import RandomForestRegressor

SheetNames = [
    "2013-14","2014-15","2015-16","2016-17","2017-18","2018-19","2020-21","2021-22","2022-23","2023-24"
]

ExcelFile = "(MBB) Nextech Stats Showdown - NCAA MENS BB Stats 2013-2024.xlsx"
AllData = []
for Sheet in SheetNames:
    DfYear = pd.read_excel(ExcelFile, sheet_name=Sheet)
    EndYearNum = int(Sheet[-2:]) + 2000
    DfYear["Year"] = EndYearNum
    TournWinsCol = f"{EndYearNum} NCAA Tournament Wins"
    DfYear.rename(
        columns={TournWinsCol: "TournWins", "Cinderella?": "Cinderella"},
        inplace=True,
        errors="ignore"
    )
    AllData.append(DfYear)

DfAll = pd.concat(AllData, ignore_index=True)

AllColumns = [
    "Team","Conference","Region","Cinderella","TournWins","Made Tournament Previous Year","Game Count","Wins","Losses","Total Points","PPG","Total Opp Points","Opp PPG","Scoring Margin","NET","NET SOS","Q1 Wins","Q1 Losses","3 Pointers Made","3 Pointers Attempted","3 Point Percentage","Free Throws Made","Free Throws Attempted","Free Throw Percentage","Rebounds","RPG","Opp Rebounds","Opp RPG","Rebound Margin","Assists","Turnovers","Assist to Turnover Ratio","TOPG","Steals","Steals per Game","Blocks","Blocks per Game","Personal Fouls","PFPG","DQs"
]

NumericCols = [
    "Made Tournament Previous Year","Cinderella","Game Count","Wins","Losses","Total Points","PPG","Total Opp Points","Opp PPG","Scoring Margin","3 Pointers Made","3 Pointers Attempted","3 Point Percentage","Free Throws Made","Free Throws Attempted","Free Throw Percentage","Rebounds","RPG","Opp Rebounds","Opp RPG","Rebound Margin","Assists","Turnovers","Assist to Turnover Ratio","TOPG","Steals","Steals per Game","Blocks","Blocks per Game","Personal Fouls","PFPG","DQs"
]

for Col in NumericCols:
    if Col in DfAll.columns:
        DfAll[Col] = pd.to_numeric(DfAll[Col], errors="coerce")
    else:
        DfAll[Col] = 0

DfAll["TournWins"] = pd.to_numeric(DfAll["TournWins"], errors="coerce")
DfAll.dropna(subset=["TournWins"], inplace=True)
DfAll.fillna(0, inplace=True)

TestYear = 2025
DfTrain = DfAll[DfAll["Year"] != TestYear]
DfTest = DfAll[DfAll["Year"] == TestYear]

XTrain = DfTrain[NumericCols]
YTrain = DfTrain["TournWins"]

Model = RandomForestRegressor(n_estimators=100, random_state=42)
Model.fit(XTrain, YTrain)

XTest = DfTest[NumericCols]
DfTest["PredictedWins"] = Model.predict(XTest)

DfTestSorted = DfTest.sort_values("PredictedWins", ascending=False)
Top10 = DfTestSorted.head(10).copy()
Top10["Rating"] = range(10, 0, -1)

print(Top10[["Team", "PredictedWins", "Cinderella", "Rating"]])