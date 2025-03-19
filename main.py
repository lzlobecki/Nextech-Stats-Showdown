import pandas as pd

DataFrame = pd.read_excel("(MBB) Nextech Stats Showdown - NCAA MENS BB Stats 2013-2025.xlsx", sheet_name="2024-25")
DataFrame.rename(columns={
    "3 Point Percentage": "3PT%",
    "Free Throw Percentage": "FT%",
    "Assist to Turnover Ratio": "AST/TO"
}, inplace=True)
DataFrame["Rating"] = (
    1.2 * DataFrame["3PT%"]
    + 1.1 * DataFrame["FT%"]
    + 0.25 * DataFrame["AST/TO"]
)
DataFrame.sort_values(by="Rating", ascending=False, inplace=True)
DataFrame.reset_index(drop=True, inplace=True)
RegionGroup = DataFrame.groupby("Region", group_keys=False)
TopTeams = RegionGroup.head(1).copy()
SecondTeams = RegionGroup.nth(1).copy()
TopTeams.sort_values(by="Rating", ascending=False, inplace=True)
SecondTeams.sort_values(by="Rating", ascending=False, inplace=True)
TopRankValues = [10, 9, 8, 7]
TopTeams["Rank"] = TopRankValues[:len(TopTeams)]
SecondRankValues = [6, 5, 4, 3]
SecondTeams["Rank"] = SecondRankValues[:len(SecondTeams)]
UsedTeams = pd.concat([TopTeams, SecondTeams], ignore_index=True)
RemainingDf = DataFrame[~DataFrame["Team"].isin(UsedTeams["Team"])].copy()
CinderellaCandidates = RemainingDf[RemainingDf["Cinderella?"] == 1.0].copy()
CinderellaCandidates.sort_values(by="Rating", ascending=False, inplace=True)
CinderellaTeams = pd.DataFrame()
if len(CinderellaCandidates) >= 2:
    CinderellaTeams = CinderellaCandidates.head(2).copy()
    CinderellaTeams["Rank"] = [2, 1]
elif len(CinderellaCandidates) == 1:
    CinderellaTeams = CinderellaCandidates.head(1).copy()
    CinderellaTeams["Rank"] = [2]
else:
    pass
FinalDf = pd.concat([TopTeams, SecondTeams, CinderellaTeams], ignore_index=True)
FinalDf.sort_values(by="Rank", ascending=False, inplace=True)
print("Final Assigned Ranks:\n")
print(FinalDf[["Team", "Conference", "Region", "Cinderella?", "Rating", "Rank"]])